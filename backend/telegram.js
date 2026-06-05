const fs = require('fs');
const path = require('path');
require('dotenv').config();

let isReady = false;

let TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN ? process.env.TELEGRAM_TOKEN.trim() : '';
let TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID ? process.env.TELEGRAM_CHAT_ID.trim() : '';
const TELEGRAM_ENABLED = String(process.env.TELEGRAM_ENABLED).trim() === 'true';

// Cooldown tracking (5 minutes)
const cooldowns = {
    anomaly: 0
};
const COOLDOWN_MS = 5 * 60 * 1000;

async function initTelegram() {
    if (!TELEGRAM_ENABLED) {
        console.log('[TELEGRAM] Telegram integration is disabled.');
        return;
    }

    if (!TELEGRAM_TOKEN || TELEGRAM_TOKEN.trim() === '') {
        console.warn('[TELEGRAM] ⚠️ TELEGRAM_TOKEN is missing in .env!');
        console.warn('[TELEGRAM] Please create a bot via @BotFather and paste the token in .env');
        return;
    }

    // Auto-detect CHAT_ID if missing
    if (!TELEGRAM_CHAT_ID || TELEGRAM_CHAT_ID.trim() === '') {
        console.log('[TELEGRAM] CHAT_ID is empty. Attempting to fetch from recent bot messages...');
        try {
            // Node 18+ has global fetch()
            const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getUpdates`);
            const data = await response.json();
            
            if (data.ok && data.result.length > 0) {
                // Get the chat ID from the latest message
                const lastMsg = data.result[data.result.length - 1];
                if (lastMsg.message && lastMsg.message.chat) {
                    TELEGRAM_CHAT_ID = lastMsg.message.chat.id.toString();
                    console.log(`[TELEGRAM] ✅ Successfully auto-detected CHAT_ID: ${TELEGRAM_CHAT_ID}`);
                    
                    // Optional: Append it to .env file for persistence
                    const envPath = path.join(__dirname, '.env');
                    let envContent = fs.readFileSync(envPath, 'utf8');
                    envContent = envContent.replace(/TELEGRAM_CHAT_ID=.*/g, `TELEGRAM_CHAT_ID=${TELEGRAM_CHAT_ID}`);
                    if (!envContent.includes('TELEGRAM_CHAT_ID=')) {
                        envContent += `\nTELEGRAM_CHAT_ID=${TELEGRAM_CHAT_ID}`;
                    }
                    fs.writeFileSync(envPath, envContent);
                    console.log('[TELEGRAM] ✅ Saved CHAT_ID to .env file!');
                    
                    isReady = true;
                    // Kirim pesan salam
                    await sendTelegramAlert('🤖 *NetVision Pro Alert Bot Aktif!*\nSistem monitoring jaringan telah berhasil dihubungkan.');
                }
            } else {
                console.warn('[TELEGRAM] ❌ Failed to detect CHAT_ID. Make sure you sent a message (e.g. /start) to the bot first!');
            }
        } catch (err) {
            console.error('[TELEGRAM] Error fetching updates:', err.message);
        }
    } else {
        console.log('[TELEGRAM] ✅ Telegram Bot is READY!');
        isReady = true;
    }
}

async function sendTelegramAlert(message) {
    if (!isReady) return;
    if (!TELEGRAM_TOKEN || !TELEGRAM_CHAT_ID) return;

    try {
        const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: message,
                parse_mode: 'Markdown'
            })
        });
        
        const data = await response.json();
        if (data.ok) {
            console.log('[TELEGRAM] ✉️  Alert sent successfully.');
        } else {
            console.error('[TELEGRAM] Failed to send alert:', data.description);
        }
    } catch (error) {
        console.error('[TELEGRAM] HTTP Error sending message:', error.message);
    }
}

async function sendAnomaly(metrics, anomalyReason) {
    if (!TELEGRAM_ENABLED || !isReady) return;

    const now = Date.now();
    if (now - cooldowns.anomaly < COOLDOWN_MS) {
        console.log('[TELEGRAM] Anomaly alert on cooldown. Skipping message.');
        return;
    }
    cooldowns.anomaly = now;

    const date = new Date();
    const timeStr = date.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const dateStr = date.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });

    const msg = `⚠️ *ANOMALI JARINGAN TERDETEKSI*

🕐 Waktu: ${timeStr} - ${dateStr}
📍 Router: ${metrics.router_ip}

📊 *Detail Anomali:*
_${anomalyReason}_

📈 *Metrik Saat Ini:*
• CPU: \`${metrics.cpu_usage}%\`
• Download: \`${metrics.rx_speed_mbps} Mbps\`
• Upload: \`${metrics.tx_speed_mbps} Mbps\`

🔗 Dashboard: http://jarkom.clyuti.my.id`;

    await sendTelegramAlert(msg);
}

async function sendNewDevice(deviceInfo, metrics) {
    if (!TELEGRAM_ENABLED || !isReady) return;

    const date = new Date();
    const timeStr = date.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const dateStr = date.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });

    const msg = `📱 *PERANGKAT BARU TERDETEKSI*

🕐 Waktu: ${timeStr} - ${dateStr}
📍 Router: ${metrics.router_ip}

💻 *Info Perangkat:*
• Nama: \`${deviceInfo.hostname || 'Unknown'}\`
• IP: \`${deviceInfo.ip || '-'}\`
• MAC: \`${deviceInfo.mac}\`
• SSID: \`${deviceInfo.ssid || '-'}\`

⚠️ _Jika bukan perangkat Anda, segera periksa jaringan WiFi Anda!_`;

    await sendTelegramAlert(msg);
}

module.exports = {
    initTelegram,
    isTelegramReady: () => isReady,
    sendAnomaly,
    sendNewDevice
};
