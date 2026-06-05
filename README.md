# NetVision Pro 🚀

NetVision Pro is an integrated, agentless local network monitoring system designed specifically for ZTE F670L routers. It automates the extraction of network telemetry data via web scraping, detects network anomalies using a dynamic Z-Score algorithm, and visualizes the results on an interactive React dashboard.

## ✨ Key Features
- **Agentless Web Scraping:** Extracts Rx/Tx bytes and CPU load directly from the router's local administration interface using Node.js and Puppeteer.
- **Dynamic Anomaly Detection:** Utilizes a multi-dimensional Z-Score algorithm with dynamic retraining to instantly detect unusual traffic spikes or high CPU loads.
- **Real-time Telegram Alerts:** Sends immediate warning notifications directly to your phone when an anomaly is detected.
- **Cloud Database Logging:** Stores historical network logs securely on Supabase (PostgreSQL).
- **Bento Grid Dashboard:** A beautiful, responsive React.js dashboard showing real-time network curves and red indicator dots for anomalies.
- **Global Accessibility:** Safely exposed to the public internet via Cloudflare Tunnel without needing router port-forwarding.

## 🛠️ Technology Stack
- **Backend:** Node.js, Express, Puppeteer
- **Frontend:** React.js, Vite, Tailwind CSS, Recharts, Lucide Icons
- **Database:** Supabase (PostgreSQL)
- **Networking:** Cloudflare Tunnel

## 📋 Prerequisites
Before running this project, ensure you have:
1. **Node.js** (v16 or higher) installed.
2. A **Supabase** account and a new project created.
3. A **Telegram Bot Token** and your Chat ID (for receiving alerts).
4. A **ZTE F670L Router** (or you can modify the Puppeteer scraper logic for your specific router model).
5. **Cloudflared** CLI installed (optional, for exposing the dashboard to a public URL).

---

## 🚀 Step-by-Step Installation

### 1. Clone the repository
```bash
git clone https://github.com/rizuca/netvision-pro.git
cd netvision-pro
```

### 2. Backend Setup
Open a terminal and navigate to the backend folder to install the required dependencies:
```bash
cd backend
npm install
```

Configure your environment variables:
- Duplicate the `.env.example` file and rename the copy to `.env`.
- Fill in your actual credentials inside `.env`:
  ```env
  PORT=5000
  SUPABASE_URL=https://your-project-id.supabase.co
  SUPABASE_KEY=your-supabase-anon-key
  ROUTER_IP=192.168.1.1
  ROUTER_USER=admin
  ROUTER_PASS=your-router-password
  MOCK_MODE=false
  ANOMALY_THRESHOLD=3.0
  TELEGRAM_TOKEN=your-telegram-bot-token
  TELEGRAM_CHAT_ID=your-telegram-chat-id
  TELEGRAM_ENABLED=true
  ```

Start the backend server (Data Collector & Anomaly Detector):
```bash
npm run dev
```

### 3. Frontend Setup
Open a **new terminal window/tab**, and navigate to the frontend folder:
```bash
cd frontend
npm install
```

Start the frontend React dashboard:
```bash
npm run dev
```
By default, the dashboard will be available at `http://localhost:5173`.

### 4. Public Access with Cloudflare Tunnel (Optional)
If you want to access your dashboard from anywhere in the world, you can run Cloudflare Tunnel:
```bash
cloudflared tunnel run your_tunnel_name
```
*Note: Ensure your Cloudflare config maps the tunnel to your local frontend port.*

---

## 🧪 Testing with Mock Mode
If you want to test the dashboard and the Telegram alerts without actually scraping a router, simply open the `backend/.env` file and set:
```env
MOCK_MODE=true
```
This will generate fake network data and artificial anomaly spikes every few minutes for demonstration purposes.

## 📄 License
This project is for educational and research purposes. Feel free to use and modify it!
