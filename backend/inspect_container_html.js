const puppeteer = require('puppeteer');
require('dotenv').config();

const ROUTER_IP   = process.env.ROUTER_IP   || '192.168.1.1';
const ROUTER_USER = process.env.ROUTER_USER  || 'admin';
const ROUTER_PASS = process.env.ROUTER_PASS  || 'admin';

(async () => {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
  });

  const page = await browser.newPage();
  try {
    await page.goto(`http://${ROUTER_IP}/`, { waitUntil: 'domcontentloaded', timeout: 15000 });
    await page.type('#Frm_Username', ROUTER_USER);
    await page.type('#Frm_Password', ROUTER_PASS);
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {}),
      page.click('#LoginId')
    ]);
    await new Promise(r => setTimeout(r, 2000));

    console.log('Clicking Local Network...');
    await page.click('#localnet');
    await new Promise(r => setTimeout(r, 1500));

    console.log('Clicking Status...');
    await page.click('#localNetStatus');
    await new Promise(r => setTimeout(r, 3000));

    console.log('Clicking WLAN Client Status Bar...');
    await page.click('#Wlan_ClientStatBar');
    await new Promise(r => setTimeout(r, 2000));

    console.log('Clicking LAN Client Status Bar...');
    await page.click('#LANDevsBar');
    await new Promise(r => setTimeout(r, 3000));

    // Print HTML structures of both containers
    const htmls = await page.evaluate(() => {
      const wlan = document.getElementById('Wlan_ClientStat_container');
      const lan = document.getElementById('LANDevs_container');
      return {
        wlan: wlan ? wlan.outerHTML : 'Not found',
        lan: lan ? lan.outerHTML : 'Not found'
      };
    });

    console.log('--- WLAN Client Status Container HTML ---');
    console.log(htmls.wlan);
    console.log('\n--- LAN Client Status Container HTML ---');
    console.log(htmls.lan);

    // Capture screenshot to visually verify
    await page.screenshot({ path: 'clients_page.png', fullPage: true });
    console.log('Screenshot saved to clients_page.png');

  } catch (err) {
    console.error('Error:', err);
  } finally {
    await browser.close();
  }
})();
