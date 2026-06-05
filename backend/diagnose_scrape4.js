/**
 * Diagnosa v4: Scrape #2 gagal menemukan elemen — artinya WLANStatusBar tidak terklik.
 * 
 * Masalah: loginAndNavigate() di scrape #2 mungkin session expired (perlu login ulang)
 * tapi setelah login, navigasi gagal karena timing.
 * 
 * Solusi: Klik WLANStatusBar lalu TUNGGU, lalu baca. Dan gunakan tombol 
 * [WLANStatus_Btn_refresh] yang ditemukan di diagnosa v3!
 * 
 * Test: Login → Navigate → Klik Refresh button → Baca → Tunggu → Klik Refresh → Baca lagi
 */
const puppeteer = require('puppeteer');
require('dotenv').config();

const ROUTER_IP   = process.env.ROUTER_IP   || '192.168.1.1';
const ROUTER_USER = process.env.ROUTER_USER  || 'admin';
const ROUTER_PASS = process.env.ROUTER_PASS  || 'admin';

async function fullScrape(page, label) {
  console.log(`\n=== ${label} ===`);
  
  // Goto root
  await page.goto(`http://${ROUTER_IP}/`, { waitUntil: 'domcontentloaded', timeout: 15000 });
  await new Promise(r => setTimeout(r, 1000));
  
  // Login jika perlu
  const isLoginPage = await page.$('input[type="password"]') !== null;
  if (isLoginPage) {
    console.log('  Login...');
    await page.type('#Frm_Username', ROUTER_USER);
    await page.type('#Frm_Password', ROUTER_PASS);
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {}),
      page.click('#LoginId')
    ]);
    await new Promise(r => setTimeout(r, 2500));
  } else {
    console.log('  Session aktif, skip login.');
  }
  
  // Navigate
  console.log('  Navigate: localnet → localNetStatus → WLANStatusBar...');
  await page.click('#localnet');
  await new Promise(r => setTimeout(r, 1500));
  await page.click('#localNetStatus');
  await new Promise(r => setTimeout(r, 1500));
  await page.click('#WLANStatusBar');
  await new Promise(r => setTimeout(r, 2500));
  
  // Coba klik tombol Refresh WLAN Status
  const refreshBtn = await page.$('#WLANStatus_Btn_refresh');
  if (refreshBtn) {
    console.log('  Klik [WLANStatus_Btn_refresh]...');
    await refreshBtn.click();
    await new Promise(r => setTimeout(r, 3000));
  } else {
    console.log('  ⚠️ Tombol WLANStatus_Btn_refresh tidak ditemukan.');
  }
  
  // Baca semua TotalBytesCount
  const data = await page.evaluate(() => {
    const results = {};
    for (let i = 0; i < 8; i++) {
      const el = document.getElementById(`TotalBytesCount:${i}`);
      if (el) {
        results[`TotalBytesCount:${i}`] = el.getAttribute('title') || el.textContent.trim();
      }
    }
    return results;
  });
  
  if (Object.keys(data).length === 0) {
    console.log('  ❌ Tidak ada TotalBytesCount ditemukan!');
  } else {
    Object.entries(data).forEach(([id, val]) => {
      if (val !== '0/0') console.log(`  ${id} = ${val}`);
    });
  }
  
  return data;
}

(async () => {
  console.log('🔍 Diagnosa v4: Test dengan Refresh button + goto ulang');
  console.log(`   Target: http://${ROUTER_IP}`);
  console.log('   ⚡ Pastikan ada traffic WiFi aktif!\n');
  
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
  });
  
  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(20000);
  
  try {
    const r1 = await fullScrape(page, 'SCRAPE #1');
    
    console.log('\n⏳ Menunggu 60 detik... (pastikan ada traffic WiFi!)');
    await new Promise(r => setTimeout(r, 60000));
    
    const r2 = await fullScrape(page, 'SCRAPE #2 (setelah 60 detik)');
    
    // Bandingkan
    console.log('\n\n========== PERBANDINGAN ==========');
    const keys = [...new Set([...Object.keys(r1), ...Object.keys(r2)])];
    let anyChanged = false;
    for (const key of keys) {
      const v1 = r1[key] || '(N/A)';
      const v2 = r2[key] || '(N/A)';
      if (v1 === '0/0' && v2 === '0/0') continue; // skip inactive
      const changed = v1 !== v2;
      if (changed) anyChanged = true;
      const icon = changed ? '🔄' : '  ';
      console.log(`  ${icon} ${key}: ${v1} → ${v2}`);
    }
    
    if (anyChanged) {
      console.log('\n✅ BERHASIL! Data berubah. Metode goto() + navigate + refresh button WORKS.');
      
      // Hitung speed untuk yang berubah
      for (const key of keys) {
        const v1 = r1[key];
        const v2 = r2[key];
        if (!v1 || !v2 || v1 === v2 || v1 === '0/0') continue;
        try {
          const [rx1, tx1] = v1.split('/').map(Number);
          const [rx2, tx2] = v2.split('/').map(Number);
          const rxMbps = (((rx2 - rx1) * 8) / (60 * 1e6)).toFixed(2);
          const txMbps = (((tx2 - tx1) * 8) / (60 * 1e6)).toFixed(2);
          console.log(`\n  ${key} speed: ↓${rxMbps} Mbps ↑${txMbps} Mbps`);
        } catch (e) {}
      }
    } else {
      console.log('\n❌ Masih sama. Router benar-benar tidak update counter.');
      console.log('   Ini mungkin karena:');
      console.log('   - Laptop terhubung via Ethernet (bukan WiFi) → counter WiFi tidak naik');
      console.log('   - Tidak ada device lain yang pakai WiFi');
      console.log('   - Counter hanya update saat ada traffic di SSID spesifik itu');
    }
    
  } catch (err) {
    console.error('\n❌ ERROR:', err.message);
  } finally {
    await browser.close();
    console.log('\n✅ Selesai.');
  }
})();
