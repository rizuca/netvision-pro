const express    = require('express');
const cors       = require('cors');
const { execSync }   = require('child_process');
const util       = require('util');
const execAsync  = util.promisify(require('child_process').exec);
const { createClient } = require('@supabase/supabase-js');
const puppeteer  = require('puppeteer');
const path       = require('path');
require('dotenv').config();
const telegram = require('./telegram');

const app  = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

// ========================
// SUPABASE SETUP
// ========================
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
let supabase = null;

if (supabaseUrl && supabaseKey && !supabaseUrl.includes('your-supabase')) {
  supabase = createClient(supabaseUrl, supabaseKey);
  console.log('✅ Supabase client initialized.');
} else {
  console.warn('⚠️  Supabase credentials not configured. DB logs will not be saved.');
}

// ========================
// CONFIGURATION
// ========================
const MOCK_MODE         = process.env.MOCK_MODE === 'true';
const ANOMALY_THRESHOLD = parseFloat(process.env.ANOMALY_THRESHOLD || '3.0');
const ROUTER_IP         = process.env.ROUTER_IP   || '192.168.1.1';
const ROUTER_USER       = process.env.ROUTER_USER  || 'admin';
const ROUTER_PASS       = process.env.ROUTER_PASS  || 'admin';

console.log(`\n🚀 Mode: ${MOCK_MODE ? '🎭 MOCK (Simulasi)' : '🌐 LIVE (Scrape Router ZTE F670L)'}`);
if (!MOCK_MODE) {
  console.log(`   Router Target : http://${ROUTER_IP}`);
  console.log(`   Router User   : ${ROUTER_USER}`);
  console.log(`   Poll Interval : 30 detik`);
}

// ========================
// IN-MEMORY STATE
// ========================
let currentState = {
  router_ip:      ROUTER_IP,
  cpu_usage:      0,
  uptime:         '0h 0m 0s',
  rx_bytes:       0n,
  tx_bytes:       0n,
  rx_speed_mbps:  0.0,
  tx_speed_mbps:  0.0,
  clients:        [],
  client_count:   0,
  is_anomaly:     false,
  anomaly_reason: null,
  is_online:      false,
  is_mock:        MOCK_MODE,
  timestamp:      Date.now()
};

let prevState = {
  rx_bytes:  0n,
  tx_bytes:  0n,
  timestamp: 0
};

const clientFirstSeen = new Map();

// ========================
// ANOMALY DETECTION MODEL (Z-Score)
// ========================
let cpuStats  = { mean: 20.0, stdDev: 10.0 };
let rxStats   = { mean: 25.0, stdDev: 15.0 };
let txStats   = { mean: 5.0,  stdDev: 3.0  };
let isTrained = false;

async function trainAnomalyDetector() {
  if (!supabase) return;
  try {
    const { data, error } = await supabase
      .from('network_logs')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(100);
    if (error) throw error;
    if (!data || data.length < 20) {
      console.log(`[TRAINING] Data kurang (${data?.length || 0}/20). Pakai threshold statis.`);
      isTrained = false;
      return;
    }
    const chronological = [...data].reverse();
    const cpus = [], rxSpeeds = [], txSpeeds = [];
    for (let i = 0; i < chronological.length; i++) {
      cpus.push(Number(chronological[i].cpu_usage));
      if (i > 0) {
        const cur  = chronological[i];
        const prev = chronological[i - 1];
        const timeDiff = (new Date(cur.created_at) - new Date(prev.created_at)) / 1000;
        if (timeDiff > 0) {
          const rxDiff = Number(BigInt(cur.rx_bytes) - BigInt(prev.rx_bytes));
          const txDiff = Number(BigInt(cur.tx_bytes) - BigInt(prev.tx_bytes));
          if (rxDiff >= 0) rxSpeeds.push((rxDiff * 8) / (timeDiff * 1e6));
          if (txDiff >= 0) txSpeeds.push((txDiff * 8) / (timeDiff * 1e6));
        }
      }
    }
    const calcStats = arr => {
      if (!arr.length) return { mean: 0, stdDev: 1 };
      const mean   = arr.reduce((a, b) => a + b, 0) / arr.length;
      const stdDev = Math.sqrt(arr.reduce((a, b) => a + (b - mean) ** 2, 0) / arr.length) || 1;
      return { mean, stdDev };
    };
    cpuStats  = calcStats(cpus);
    rxStats   = calcStats(rxSpeeds);
    txStats   = calcStats(txSpeeds);
    isTrained = true;
    console.log(`[TRAINING] ✅ Model dilatih dari ${data.length} record.`);
    console.log(`  CPU: μ=${cpuStats.mean.toFixed(1)}% σ=${cpuStats.stdDev.toFixed(1)}%`);
    console.log(`  Rx:  μ=${rxStats.mean.toFixed(2)} Mbps σ=${rxStats.stdDev.toFixed(2)} Mbps`);
    console.log(`  Tx:  μ=${txStats.mean.toFixed(2)} Mbps σ=${txStats.stdDev.toFixed(2)} Mbps`);
  } catch (err) {
    console.error('[TRAINING ERROR]', err.message);
  }
}

// ========================
// ANOMALY DETECTION LOGIC
// ========================
function detectAnomalies(cpu, rxMbps, txMbps) {
  const reasons = [];
  if (isTrained) {
    const zCpu = Math.abs(cpu    - cpuStats.mean) / cpuStats.stdDev;
    const zRx  = Math.abs(rxMbps - rxStats.mean)  / rxStats.stdDev;
    const zTx  = Math.abs(txMbps - txStats.mean)   / txStats.stdDev;
    if (zCpu > ANOMALY_THRESHOLD && cpu    > 60)  reasons.push(`CPU Load abnormal: ${cpu}% (normal ~${cpuStats.mean.toFixed(0)}%)`);
    if (zRx  > ANOMALY_THRESHOLD && rxMbps > 50)  reasons.push(`Download abnormal: ${rxMbps} Mbps (normal ~${rxStats.mean.toFixed(0)} Mbps)`);
    if (zTx  > ANOMALY_THRESHOLD && txMbps > 20)  reasons.push(`Upload abnormal: ${txMbps} Mbps (normal ~${txStats.mean.toFixed(0)} Mbps)`);
  } else {
    if (cpu    > 85)  reasons.push(`CPU Load sangat tinggi: ${cpu}%`);
    if (rxMbps > 150) reasons.push(`Download sangat tinggi: ${rxMbps} Mbps`);
    if (txMbps > 50)  reasons.push(`Upload sangat tinggi: ${txMbps} Mbps`);
  }
  return reasons;
}

// ========================
// SYSTEM STATS (CPU + Uptime + Ping)
// ========================
async function getPingLatency() {
  try {
    const { stdout } = await execAsync('ping -n 1 8.8.8.8');
    const match = stdout.match(/time[=<](\d+)ms/i);
    if (match) return parseInt(match[1], 10);
  } catch (e) {
    // Ping gagal (RTO/Offline)
  }
  return 0; // default 0 jika gagal
}

function getWindowsUptimeSeconds() {
  try {
    const raw = execSync(
      `powershell -Command "(Get-Date) - (gcim Win32_OperatingSystem).LastBootUpTime | Select-Object -ExpandProperty TotalSeconds"`,
      { timeout: 3000 }
    ).toString().trim();
    return Math.floor(parseFloat(raw));
  } catch {
    return Math.floor(process.uptime());
  }
}

function getWindowsCpuPercent() {
  try {
    const raw = execSync(
      `powershell -Command "Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average"`,
      { timeout: 3000 }
    ).toString().trim();
    return Math.round(parseFloat(raw));
  } catch {
    return Math.floor(10 + Math.random() * 15);
  }
}

function formatUptime(totalSeconds) {
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return `${h}h ${m}m ${s}s`;
}

// ========================
// PUPPETEER BROWSER + PAGE SINGLETON
// ========================
let browserInstance = null;
let persistentPage = null;

async function getBrowserInstance() {
  if (browserInstance) {
    try {
      await browserInstance.pages();
      return browserInstance;
    } catch (e) {
      console.warn('[BROWSER] Crash terdeteksi. Meluncurkan ulang browser...');
      browserInstance = null;
      persistentPage = null;
      isPageReady = false;
    }
  }

  console.log('[BROWSER] Meluncurkan Puppeteer (headless Chrome)...');
  browserInstance = await puppeteer.launch({
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--no-first-run',
    ]
  });
  console.log('[BROWSER] ✅ Browser siap digunakan.');
  return browserInstance;
}

async function getPersistentPage() {
  const browser = await getBrowserInstance();

  // Cek apakah persistent page masih hidup
  if (persistentPage) {
    try {
      await persistentPage.title(); // throws jika page sudah tertutup/crash
      return persistentPage;
    } catch (e) {
      console.warn('[PAGE] Page mati. Membuat page baru...');
      persistentPage = null;
    }
  }

  console.log('[PAGE] Membuat persistent page baru...');
  persistentPage = await browser.newPage();
  persistentPage.setDefaultNavigationTimeout(20000);
  persistentPage.on('console', () => {});
  persistentPage.on('pageerror', () => {});
  return persistentPage;
}

// ========================
// ROUTER LOGIN
// ========================
async function loginToRouter(page) {
  await page.goto(`http://${ROUTER_IP}/`, { waitUntil: 'domcontentloaded', timeout: 15000 });

  // Detect if login form is present
  const isLoginPage = await page.$('input[type="password"]') !== null;
  if (!isLoginPage) {
    console.log('[ROUTER] Sesi masih aktif, melewati login.');
    return;
  }

  console.log('[ROUTER] Halaman login terdeteksi. Mengisi kredensial...');

  // Username — coba multiple selector (firmware ZTE berbeda-beda)
  const userSelectors = [
    '#Frm_Username', '#txt_Username', '#UserName', '#username',
    'input[name="Username"]', 'input[name="username"]',
    'input[type="text"]'
  ];
  for (const sel of userSelectors) {
    try {
      const el = await page.$(sel);
      if (el) {
        await el.click({ clickCount: 3 }); // select all
        await el.type(ROUTER_USER);
        break;
      }
    } catch (e) { /* try next */ }
  }

  // Password
  const passSelectors = [
    '#Frm_Password', '#txt_Password', '#Password', '#password',
    'input[name="Password"]', 'input[name="password"]',
    'input[type="password"]'
  ];
  for (const sel of passSelectors) {
    try {
      const el = await page.$(sel);
      if (el) {
        await el.click({ clickCount: 3 });
        await el.type(ROUTER_PASS);
        break;
      }
    } catch (e) { /* try next */ }
  }

  // Submit button
  const btnSelectors = [
    '#LoginId', '#Btn_Login', '#btnLogin', '#loginBtn',
    'input[type="submit"]', 'button[type="submit"]',
    'button'
  ];
  for (const sel of btnSelectors) {
    try {
      const el = await page.$(sel);
      if (el) {
        await Promise.all([
          page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {}),
          el.click()
        ]);
        break;
      }
    } catch (e) { /* try next */ }
  }

  // Wait for page to settle after login
  await new Promise(r => setTimeout(r, 2500));

  // Verify login success
  const stillOnLogin = await page.$('input[type="password"]') !== null;
  if (stillOnLogin) {
    throw new Error('Login gagal! Periksa ROUTER_USER dan ROUTER_PASS di file .env');
  }
  console.log('[ROUTER] ✅ Login berhasil!');
}

// ========================
// NAVIGATE TO WLAN STATUS PAGE (only if not already there)
// ========================
async function ensureOnWlanStatusPage(page) {
  // Setiap polling HARUS goto() ulang karena router ZTE tidak update DOM in-place
  try {
    console.log('[PAGE] goto() → login → navigate → refresh...');
    await page.goto(`http://${ROUTER_IP}/`, { waitUntil: 'domcontentloaded', timeout: 15000 });
    await new Promise(r => setTimeout(r, 1000));

    // Login jika session expired
    const isLoginPage = await page.$('input[type="password"]') !== null;
    if (isLoginPage) {
      await loginToRouter(page);
    }

    // Navigate: Local Network → Status → WLAN Status
    await page.waitForSelector('#localnet', { timeout: 10000 });
    await page.click('#localnet');
    await new Promise(r => setTimeout(r, 1500));

    await page.waitForSelector('#localNetStatus', { timeout: 10000 });
    await page.click('#localNetStatus');
    await new Promise(r => setTimeout(r, 1500));

    await page.waitForSelector('#WLANStatusBar', { timeout: 10000 });
    await page.click('#WLANStatusBar');
    await new Promise(r => setTimeout(r, 2000));

    // Klik tombol Refresh untuk paksa router update counter
    const refreshBtn = await page.$('#WLANStatus_Btn_refresh');
    if (refreshBtn) {
      await refreshBtn.click();
      await new Promise(r => setTimeout(r, 3000));
    }

    // Expand Client Status panel
    const clientBar = await page.$('#Wlan_ClientStatBar');
    if (clientBar) {
      await clientBar.click();
      await new Promise(r => setTimeout(r, 1500));
      // Klik refresh client juga
      const clientRefresh = await page.$('#Btn_refresh_Wlan_ClientStat');
      if (clientRefresh) {
        await clientRefresh.click();
        await new Promise(r => setTimeout(r, 2000));
      }
    }

    // Verifikasi
    const el = await page.$('[id="TotalBytesCount:0"]');
    if (el) {
      return true;
    }
  } catch (e) {
    console.error('[PAGE NAV ERROR]', e.message);
  }

  return false;
}

// ========================
// PUPPETEER ROUTER SCRAPER (Page Singleton)
// ========================
async function scrapeRouterStats() {
  try {
    const page = await getPersistentPage();

    const success = await ensureOnWlanStatusPage(page);
    if (!success) {
      throw new Error('Tidak dapat menavigasi ke halaman WLAN Status.');
    }

    // Extract bytes from ALL TotalBytesCount (sum semua SSID: 2.4G + 5G)
    const bytesData = await page.evaluate(() => {
      let totalRx = 0;
      let totalTx = 0;
      const details = [];
      
      for (let i = 0; i < 8; i++) {
        const el = document.getElementById(`TotalBytesCount:${i}`);
        if (!el) continue;
        const raw = el.getAttribute('title') || el.textContent.trim();
        if (!raw || raw === '0/0') continue;
        
        const cleaned = raw.replace(/[^\d/]/g, '');
        const parts = cleaned.split('/');
        if (parts.length === 2 && parts[0] && parts[1]) {
          const rx = parseInt(parts[0], 10);
          const tx = parseInt(parts[1], 10);
          if (!isNaN(rx) && !isNaN(tx)) {
            totalRx += rx;
            totalTx += tx;
            details.push({ index: i, rx, tx });
          }
        }
      }
      
      return { totalRx, totalTx, details };
    });

    if (bytesData.totalRx === 0 && bytesData.totalTx === 0) {
      throw new Error('Semua TotalBytesCount bernilai 0 — tidak ada traffic terdeteksi.');
    }

    const rxBytes = BigInt(bytesData.totalRx);
    const txBytes = BigInt(bytesData.totalTx);

    console.log(`[ROUTER] Scrape sukses → Rx: ${rxBytes} bytes | Tx: ${txBytes} bytes (dari ${bytesData.details.length} SSID aktif)`);
    bytesData.details.forEach(d => console.log(`         SSID:${d.index} → Rx:${d.rx} Tx:${d.tx}`));

    // Extract connected clients
    const clients = await page.evaluate(() => {
      const list = [];
      const divs = Array.from(document.querySelectorAll('[id^="template_Wlan_ClientStat_"]'));
      for (const div of divs) {
        if (div.style.display === 'none') continue;
        
        const idParts = div.id.split('_');
        const idx = idParts[idParts.length - 1];
        if (isNaN(idx)) continue;
        
        const hostEl = document.getElementById(`HostName:${idx}`);
        const ipEl   = document.getElementById(`IPAddress:${idx}`);
        const macEl  = document.getElementById(`MACAddress:${idx}`);
        const ssidEl = document.getElementById(`ESSID:${idx}`);
        
        const hostname = hostEl ? (hostEl.getAttribute('title') || hostEl.textContent.trim()) : 'Unknown';
        const ip       = ipEl ? (ipEl.getAttribute('title') || ipEl.textContent.trim()) : '';
        const mac      = macEl ? (macEl.getAttribute('title') || macEl.textContent.trim()) : '';
        const ssid     = ssidEl ? (ssidEl.getAttribute('title') || ssidEl.textContent.trim()) : '';
        
        if (mac) {
          list.push({ hostname, ip, mac, ssid });
        }
      }
      return list;
    });

    console.log(`[ROUTER] Clients: ${clients.length}`);
    return { rxBytes, txBytes, clients };

  } catch (err) {
    console.error('[ROUTER SCRAPE ERROR]', err.message);
    // Reset page singleton on error
    persistentPage = null;
    if (browserInstance) {
      try { await browserInstance.close(); } catch (e) {}
      browserInstance = null;
    }
    return null;
  }
}

// ========================
// LIVE POLLING (Router Scrape + Windows CPU + Anomaly Detection)
// ========================
async function pollLiveData() {
  const now = Date.now();

  // 1. Scrape router ZTE F670L untuk Rx/Tx bytes
  const routerStats = await scrapeRouterStats();
  if (!routerStats) {
    currentState.is_online = false;
    console.warn('[POLL] Scraping router gagal. Data tidak diperbarui. Menunggu siklus berikutnya (30s)...');
    return;
  }

  // 2. Hitung kecepatan bandwidth (Mbps) dari selisih bytes
  if (prevState.timestamp > 0) {
    const timeDiffSec = (now - prevState.timestamp) / 1000;
    if (timeDiffSec > 0) {
      let rxDiff = routerStats.rxBytes - prevState.rx_bytes;
      let txDiff = routerStats.txBytes - prevState.tx_bytes;
      // Guard counter reset (router reboot)
      if (rxDiff < 0n) rxDiff = 0n;
      if (txDiff < 0n) txDiff = 0n;
      currentState.rx_speed_mbps = parseFloat(((Number(rxDiff) * 8) / (timeDiffSec * 1e6)).toFixed(2));
      currentState.tx_speed_mbps = parseFloat(((Number(txDiff) * 8) / (timeDiffSec * 1e6)).toFixed(2));
    }
  } else {
    // Siklus pertama — belum ada data sebelumnya, kecepatan = 0
    currentState.rx_speed_mbps = 0.0;
    currentState.tx_speed_mbps = 0.0;
  }

  // 3. Ambil CPU, Uptime & Ping
  const cpu       = getWindowsCpuPercent();
  const uptimeSec = getWindowsUptimeSeconds();
  const ping_ms   = await getPingLatency();

  // 4. Update state global
  currentState.cpu_usage    = cpu;
  currentState.uptime       = formatUptime(uptimeSec);
  currentState.ping_ms      = ping_ms;
  currentState.rx_bytes     = routerStats.rxBytes;
  currentState.tx_bytes     = routerStats.txBytes;

  const currentClients = routerStats.clients || [];
  currentClients.forEach(c => {
    if (c.mac && !clientFirstSeen.has(c.mac)) {
      clientFirstSeen.set(c.mac, Date.now());
      telegram.sendNewDevice(c, currentState);
    }
    c.active_since = clientFirstSeen.get(c.mac) || Date.now();
  });

  currentState.clients      = currentClients;
  currentState.client_count = currentState.clients.length;
  currentState.is_online    = true;
  currentState.is_mock      = false;
  currentState.timestamp    = now;

  // 5. Deteksi anomali Z-Score
  const reasons = detectAnomalies(cpu, currentState.rx_speed_mbps, currentState.tx_speed_mbps);
  if (reasons.length > 0) {
    currentState.is_anomaly     = true;
    currentState.anomaly_reason = reasons.join(' & ');
    console.warn(`[⚠ ANOMALI] ${currentState.anomaly_reason}`);
  } else {
    currentState.is_anomaly     = false;
    currentState.anomaly_reason = null;
  }

  prevState = { rx_bytes: routerStats.rxBytes, tx_bytes: routerStats.txBytes, timestamp: now };
  console.log(`[LIVE] CPU:${cpu}% ↓${currentState.rx_speed_mbps}Mbps ↑${currentState.tx_speed_mbps}Mbps Uptime:${currentState.uptime} Clients:${currentState.client_count}`);
}

// ========================
// MOCK DATA GENERATOR
// ========================
let pollCounter = 0;
function generateMockData() {
  const now         = Date.now();
  const timeDiffSec = prevState.timestamp ? (now - prevState.timestamp) / 1000 : 60;
  pollCounter++;

  // Interval 60s → anomaly setiap 2x poll = 120 detik ≈ 2 menit
  const triggerAnomaly = (pollCounter % 2 === 0);

  if (triggerAnomaly) {
    const type = Math.random() > 0.5 ? 'cpu' : 'bandwidth';
    if (type === 'cpu') {
      currentState.cpu_usage      = parseFloat((85 + Math.random() * 10).toFixed(1));
      currentState.rx_speed_mbps  = parseFloat((15 + Math.random() * 5).toFixed(2));
      currentState.tx_speed_mbps  = parseFloat((3  + Math.random() * 2).toFixed(2));
      currentState.ping_ms        = Math.floor(150 + Math.random() * 100);
      currentState.is_anomaly     = true;
      currentState.anomaly_reason = `Lonjakan CPU Kritis: ${currentState.cpu_usage}% (normal ~20%)`;
    } else {
      currentState.cpu_usage      = parseFloat((20 + Math.random() * 5).toFixed(1));
      currentState.rx_speed_mbps  = parseFloat((380 + Math.random() * 50).toFixed(2));
      currentState.tx_speed_mbps  = parseFloat((85  + Math.random() * 15).toFixed(2));
      currentState.ping_ms        = Math.floor(150 + Math.random() * 100);
      currentState.is_anomaly     = true;
      currentState.anomaly_reason = `Trafik Download Abnormal: ${currentState.rx_speed_mbps} Mbps (normal ~25 Mbps)`;
    }
    console.log(`[MOCK ANOMALI] ${currentState.anomaly_reason}`);
  } else {
    currentState.cpu_usage      = parseFloat((15 + Math.sin(now / 50000) * 10 + Math.random() * 5).toFixed(1));
    currentState.rx_speed_mbps  = parseFloat((25 + Math.sin(now / 30000) * 15 + Math.random() * 5).toFixed(2));
    currentState.tx_speed_mbps  = parseFloat((5  + Math.cos(now / 40000) * 4  + Math.random() * 2).toFixed(2));
    currentState.ping_ms        = Math.floor(15 + Math.random() * 25);
    currentState.is_anomaly     = false;
    currentState.anomaly_reason = null;
  }

  const uptimeSec = Math.floor(now / 1000) % 86400;
  currentState.uptime   = formatUptime(uptimeSec);
  const rxDiff = BigInt(Math.floor((currentState.rx_speed_mbps * 1e6 * timeDiffSec) / 8));
  const txDiff = BigInt(Math.floor((currentState.tx_speed_mbps * 1e6 * timeDiffSec) / 8));
  currentState.rx_bytes  = prevState.rx_bytes + rxDiff;
  currentState.tx_bytes  = prevState.tx_bytes + txDiff;

  // Generate mock clients
  currentState.clients = [
    { hostname: 'ASUS-ROG-Laptop', ip: '192.168.1.12', mac: 'b4:2e:99:ab:cd:ef', ssid: 'Winaila 5G' },
    { hostname: 'iPhone-15-Pro', ip: '192.168.1.15', mac: '0e:56:20:cc:4b:bd', ssid: 'Winaila 5G' },
    { hostname: 'Samsung-Smart-TV', ip: '192.168.1.18', mac: 'a4:70:d6:12:34:56', ssid: 'Winaila 2.4G' }
  ];
  if (currentState.is_anomaly && currentState.anomaly_reason.includes('Download')) {
    currentState.clients.push({ hostname: 'Unknown-Downloader-PC', ip: '192.168.1.99', mac: 'ff:ff:ff:ff:ff:ff', ssid: 'Winaila 5G' });
  }
  
  currentState.clients.forEach(c => {
    if (c.mac && !clientFirstSeen.has(c.mac)) {
      clientFirstSeen.set(c.mac, Date.now());
      telegram.sendNewDevice(c, currentState);
    }
    c.active_since = clientFirstSeen.get(c.mac) || Date.now();
  });

  currentState.client_count = currentState.clients.length;

  currentState.is_online = true;
  currentState.is_mock   = true;
  currentState.timestamp = now;
  prevState = { rx_bytes: currentState.rx_bytes, tx_bytes: currentState.tx_bytes, timestamp: now };
}

// ========================
// DATABASE SYNC (Supabase)
// ========================
async function saveToDatabase() {
  if (!supabase) return;
  if (!currentState.is_online && !MOCK_MODE) return;

  const payload = {
    router_ip:      currentState.router_ip,
    cpu_usage:      currentState.cpu_usage,
    uptime:         currentState.uptime,
    rx_bytes:       currentState.rx_bytes.toString(),
    tx_bytes:       currentState.tx_bytes.toString(),
    client_count:   currentState.client_count,
    is_anomaly:     currentState.is_anomaly,
    anomaly_reason: currentState.anomaly_reason
  };

  try {
    const { error } = await supabase.from('network_logs').insert([payload]);
    if (error) throw error;
    console.log('[DB] ✅ Data tersimpan ke Supabase.');
  } catch (err) {
    console.error('[DB ERROR]', err.message);
  }
}

// ========================
// MAIN POLL + SAVE CYCLE (every 60 seconds)
// ========================
async function pollAndSave() {
  if (MOCK_MODE) {
    generateMockData();
  } else {
    await pollLiveData(); // async Puppeteer scrape
  }
  
  if (currentState.is_anomaly && currentState.anomaly_reason) {
    telegram.sendAnomaly(currentState, currentState.anomaly_reason);
  }

  await saveToDatabase();
}

// ========================
// STARTUP & INTERVALS
// ========================
trainAnomalyDetector();
setInterval(trainAnomalyDetector, 300000); // Re-train Z-Score model every 5 minutes
telegram.initTelegram();

// Initial poll on startup (with delay to let server fully boot first)
setTimeout(() => {
  console.log('\n[STARTUP] Memulai polling pertama...');
  pollAndSave();
}, 3000);

// Poll & save every 30 seconds
setInterval(pollAndSave, 30000);

// ========================
// API ENDPOINTS
// ========================
app.get('/api/metrics/current', (req, res) => {
  res.json({
    ...currentState,
    rx_bytes: currentState.rx_bytes.toString(),
    tx_bytes: currentState.tx_bytes.toString()
  });
});

app.get('/api/metrics/history', async (req, res) => {
  if (!supabase) return res.status(503).json({ error: 'Supabase not configured' });
  try {
    const { data, error } = await supabase
      .from('network_logs')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(50);
    if (error) throw error;

    const sorted = [...data].reverse();
    const withSpeeds = sorted.map((item, idx) => {
      let rx_speed = 0, tx_speed = 0;
      if (idx > 0) {
        const prev = sorted[idx - 1];
        const diff = (new Date(item.created_at) - new Date(prev.created_at)) / 1000;
        if (diff > 0) {
          const rxD = Number(BigInt(item.rx_bytes) - BigInt(prev.rx_bytes));
          const txD = Number(BigInt(item.tx_bytes) - BigInt(prev.tx_bytes));
          if (rxD >= 0) rx_speed = parseFloat(((rxD * 8) / (diff * 1e6)).toFixed(2));
          if (txD >= 0) tx_speed = parseFloat(((txD * 8) / (diff * 1e6)).toFixed(2));
        }
      }
      return { ...item, rx_speed_mbps: rx_speed, tx_speed_mbps: tx_speed };
    });
    res.json(withSpeeds);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Health check endpoint (berguna untuk verifikasi Cloudflare tunnel)
app.get('/api/health', (req, res) => {
  res.json({
    status:    'ok',
    mode:      MOCK_MODE ? 'mock' : 'live',
    router_ip: ROUTER_IP,
    uptime:    currentState.uptime,
    timestamp: new Date().toISOString()
  });
});

// Model stats endpoint — ekspos statistik Z-Score untuk presentasi
app.get('/api/model/stats', (req, res) => {
  res.json({
    is_trained: isTrained,
    threshold:  ANOMALY_THRESHOLD,
    cpu:  { mean: cpuStats.mean, stdDev: cpuStats.stdDev },
    rx:   { mean: rxStats.mean,  stdDev: rxStats.stdDev  },
    tx:   { mean: txStats.mean,  stdDev: txStats.stdDev  },
    current_z_scores: {
      cpu: cpuStats.stdDev > 0 ? parseFloat((Math.abs(currentState.cpu_usage - cpuStats.mean) / cpuStats.stdDev).toFixed(2)) : 0,
      rx:  rxStats.stdDev  > 0 ? parseFloat((Math.abs(currentState.rx_speed_mbps - rxStats.mean) / rxStats.stdDev).toFixed(2)) : 0,
      tx:  txStats.stdDev  > 0 ? parseFloat((Math.abs(currentState.tx_speed_mbps - txStats.mean) / txStats.stdDev).toFixed(2)) : 0,
    },
    is_anomaly:     currentState.is_anomaly,
    anomaly_reason: currentState.anomaly_reason
  });
});

app.get('/api/telegram/status', (req, res) => {
  res.json({ ready: telegram.isTelegramReady() });
});

// Endpoint untuk testing manual
app.get('/api/telegram/test', async (req, res) => {
  if (!telegram.isTelegramReady()) {
    return res.status(400).json({ error: 'Telegram belum siap' });
  }
  
  // Test 1: Kirim Test Anomali
  await telegram.sendAnomaly(
    { 
      router_ip: currentState.router_ip || '192.168.1.1', 
      cpu_usage: 95, 
      rx_speed_mbps: 150.5, 
      tx_speed_mbps: 20.1 
    }, 
    'Trafik Download Sangat Tinggi (Test Manual)'
  );

  // Test 2: Kirim Test Perangkat Baru
  await telegram.sendNewDevice(
    { 
      hostname: 'Device-Test-Manual', 
      ip: '192.168.1.250', 
      mac: 'AA:BB:CC:DD:EE:FF', 
      ssid: 'WIFI-TEST' 
    },
    currentState
  );

  res.json({ message: 'Pesan tes anomali dan perangkat baru telah dikirim ke Telegram!' });
});

// Serve frontend static files
app.use(express.static(path.join(__dirname, '../frontend/dist')));
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/dist/index.html'));
});

// ========================
// START SERVER
// ========================
app.listen(PORT, () => {
  console.log(`\n🚀 Backend berjalan di http://localhost:${PORT}`);
  if (MOCK_MODE) {
    console.log('📊 Mode: SIMULASI — anomali muncul otomatis setiap ~2 menit.');
  } else {
    console.log(`🌐 Mode: LIVE — scraping ZTE F670L di http://${ROUTER_IP}`);
    console.log('   Sumber data: Bandwidth jaringan dari router + CPU + Uptime Windows');
    console.log('   Polling: setiap 30 detik (headless Chrome via Puppeteer)');
  }
  console.log(`🏥 Health check: http://localhost:${PORT}/api/health\n`);
});
