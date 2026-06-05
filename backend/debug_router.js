// Debug final: coba openLink() langsung untuk berbagai tag WLAN Status
const puppeteer = require('puppeteer');
require('dotenv').config();

const ROUTER_IP   = process.env.ROUTER_IP   || '192.168.1.1';
const ROUTER_USER = process.env.ROUTER_USER  || 'admin';
const ROUTER_PASS = process.env.ROUTER_PASS  || 'admin';

async function login(page) {
  await page.goto(`http://${ROUTER_IP}/`, { waitUntil: 'domcontentloaded', timeout: 15000 });
  await page.type('#Frm_Username', ROUTER_USER);
  await page.type('#Frm_Password', ROUTER_PASS);
  await Promise.all([
    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {}),
    page.click('#LoginId')
  ]);
  await new Promise(r => setTimeout(r, 2000));
}

(async () => {
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
  });

  const page = await browser.newPage();

  // ===== STRATEGI 1: Coba openLink() dengan berbagai tag =====
  const candidateTags = [
    'wlanStatus',
    'wlanstatus',
    'local_wlanStatus',
    'localNetWlanStatus',
    'wlan_status',
    'wlanStatusInfo',
    'wlanDevStatus',
    'localNetStatus_wlan',
    'wlan_wlanstatus_lua.lua',
  ];

  try {
    await login(page);
    console.log('✅ Login berhasil\n');

    console.log('=== STRATEGI 1: Coba openLink() untuk setiap kandidat tag ===\n');
    for (const tag of candidateTags) {
      // Reset ke halaman utama dulu
      const result = await page.evaluate((t) => {
        if (typeof openLink === 'function') {
          openLink(t + '&Menu3Location=0');
          return 'called';
        }
        return 'openLink not found';
      }, tag);
      
      await new Promise(r => setTimeout(r, 2000));
      
      const found = await page.$('[id="TotalBytesCount:0"]');
      if (found) {
        const val = await page.evaluate(el => el.getAttribute('title') || el.textContent, found);
        console.log(`✅✅✅ SUKSES! openLink("${tag}") → TotalBytesCount:0 = "${val}"`);
        await browser.close();
        return;
      } else {
        const title = await page.title();
        console.log(`❌ openLink("${tag}") → tidak ada (title: "${title}")`);
      }
    }

    // ===== STRATEGI 2: Klik Status submenu lalu cari sub-sub-menu =====
    console.log('\n=== STRATEGI 2: Status → sub-sub-menu ===\n');
    await page.click('#localnet');
    await new Promise(r => setTimeout(r, 1500));
    await page.click('#localNetStatus');
    await new Promise(r => setTimeout(r, 2000));

    // Tangkap semua elemen navigasi yang muncul setelah klik Status
    const statusSubItems = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('[menupage], a[onclick], li a'))
        .map(el => ({
          id:       el.id || '',
          menupage: el.getAttribute('menupage') || '',
          onclick:  el.getAttribute('onclick') || '',
          text:     el.textContent.trim().replace(/\s+/g, ' ').substring(0, 50),
          class:    el.className || '',
          style:    el.getAttribute('style') || '',
        }))
        .filter(el => el.text.length > 0);
    });

    console.log(`Sub-items setelah klik Status (${statusSubItems.length}):`);
    statusSubItems.forEach(el => {
      if (el.menupage || el.onclick) {
        console.log(`  [${el.id}] mp="${el.menupage}" onclick="${el.onclick.substring(0,60)}" → "${el.text}"`);
      }
    });

    // Klik setiap sub-item dan cek TotalBytesCount:0
    for (const item of statusSubItems) {
      if (!item.menupage && !item.onclick) continue;
      
      await page.evaluate((id, mp, oc) => {
        if (id) {
          const el = document.getElementById(id);
          if (el) { el.click(); return; }
        }
        if (mp) {
          const el = document.querySelector(`[menupage="${mp}"]`);
          if (el) { el.click(); return; }
        }
      }, item.id, item.menupage, item.onclick);
      
      await new Promise(r => setTimeout(r, 2000));
      
      const found = await page.$('[id="TotalBytesCount:0"]');
      if (found) {
        const val = await page.evaluate(el => el.getAttribute('title') || el.textContent, found);
        console.log(`\n✅✅✅ DITEMUKAN! Klik [${item.id}] menupage="${item.menupage}" → "${item.text}"`);
        console.log(`   TotalBytesCount:0 = "${val}"`);
        await browser.close();
        return;
      }
    }

    // ===== STRATEGI 3: Direct fetch AJAX =====
    console.log('\n=== STRATEGI 3: Direct AJAX fetch ===\n');
    const ajaxResults = await page.evaluate(async (ip) => {
      const tags = ['wlanStatus', 'local_wlanStatus', 'wlan_status', 'wlanDevStatus', 'wlanStatusInfo'];
      const results = [];
      for (const tag of tags) {
        try {
          const r = await fetch(`http://${ip}/?_type=menuData&_tag=${tag}&_=${Date.now()}`);
          const text = await r.text();
          results.push({ tag, status: r.status, has_bytes: text.includes('TotalBytes') || text.includes('BytesCount'), preview: text.substring(0, 100) });
        } catch (e) {
          results.push({ tag, error: e.message });
        }
      }
      return results;
    }, ROUTER_IP);

    ajaxResults.forEach(r => {
      console.log(`  _tag=${r.tag}: status=${r.status} has_bytes=${r.has_bytes}`);
      if (r.has_bytes) console.log(`    *** PREVIEW: ${r.preview}`);
    });

  } catch (err) {
    console.error(`\n❌ ERROR: ${err.message}`);
    console.error(err.stack);
  }

  await browser.close();
  console.log('\n✅ Selesai.');
})();
