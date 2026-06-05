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

    // Let's print the client containers texts first to see what's loaded
    const initialText = await page.evaluate(() => {
      const wlanContainer = document.getElementById('Wlan_ClientStat_container');
      const lanContainer = document.getElementById('LANDevs_container');
      return {
        wlan: wlanContainer ? wlanContainer.innerText : 'WLAN not found',
        lan: lanContainer ? lanContainer.innerText : 'LAN not found'
      };
    });

    console.log('Initial WLAN Container text:', initialText.wlan);
    console.log('Initial LAN Container text:', initialText.lan);

    // Let's see if there are headers we need to click to expand them
    const bars = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('h1, h2, h3, .collapBarWithDataTrans'))
        .map(el => ({ id: el.id, text: el.innerText }));
    });
    console.log('Headers on page:', bars);

    // If there are bars like WLANClientStatBar or LANDevsBar, click them
    await page.evaluate(() => {
      // Find and click headers containing client info
      const elements = Array.from(document.querySelectorAll('h1, h2, h3, .collapBarWithDataTrans'));
      for (const el of elements) {
        if (el.innerText.includes('Client') || el.id.includes('Client') || el.id.includes('Devs')) {
          console.log('Clicking:', el.id, el.innerText);
          el.click();
        }
      }
    });

    await new Promise(r => setTimeout(r, 5000));

    // Fetch the loaded elements
    const loadedData = await page.evaluate(() => {
      const extractTable = (containerId) => {
        const container = document.getElementById(containerId);
        if (!container) return 'Container not found';
        
        // Find all divs or rows under container
        const rows = [];
        const childDivs = Array.from(container.children);
        for (const child of childDivs) {
          if (child.id && child.id.startsWith('template_') && child.style.display !== 'none') {
            // This is an active data row
            const textContent = child.innerText.trim().replace(/\s+/g, ' ');
            rows.push(textContent);
          }
        }
        
        // Also look for standard tables if any
        const tables = Array.from(container.querySelectorAll('table'));
        const tableRows = [];
        tables.forEach(table => {
          Array.from(table.querySelectorAll('tr')).forEach(tr => {
            tableRows.push(tr.innerText.trim().replace(/\s+/g, ' '));
          });
        });
        
        return {
          innerHtml: container.innerHTML.substring(0, 1000),
          rows: rows,
          tableRows: tableRows
        };
      };

      return {
        wlan: extractTable('Wlan_ClientStat_container'),
        lan: extractTable('LANDevs_container')
      };
    });

    console.log('Loaded WLAN Clients:', loadedData.wlan.rows);
    console.log('Loaded LAN/WLAN Devs:', loadedData.lan.rows);

  } catch (err) {
    console.error('Error:', err);
  } finally {
    await browser.close();
  }
})();
