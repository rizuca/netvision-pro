import React, { useState, useEffect, useMemo } from 'react';
import axios from 'axios';
import { 
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { 
  Activity, 
  ArrowDown, 
  ArrowUp, 
  Cpu, 
  Clock, 
  Server, 
  Wifi, 
  AlertCircle, 
  ShieldCheck, 
  RefreshCw,
  Info,
  History,
  Users,
  Send
} from 'lucide-react';

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5000';
const ANOMALY_THRESHOLD = 3.0;

// Colors for Charts - adjusted for light mode visibility
const COLORS = {
  download: '#0891b2', // Cyan-600
  upload: '#c026d3',   // Fuchsia-600
  cpuUsed: '#2563eb',  // Blue-600
  cpuFree: 'rgba(148, 163, 184, 0.15)',
  anomaly: '#dc2626'   // Red-600
};

// Bento Card Component — Light Glassmorphism
const BentoCard = ({ children, className = '', title, icon: Icon }) => (
  <div className={`bg-white/70 backdrop-blur-xl border border-slate-200/80 rounded-3xl p-6 shadow-[0_4px_20px_0_rgba(100,116,139,0.10)] flex flex-col relative overflow-hidden transition-all duration-300 hover:border-indigo-200 hover:bg-white/90 hover:shadow-[0_8px_30px_0_rgba(99,102,241,0.10)] hover:-translate-y-0.5 ${className}`}>
    {/* Subtle card top-right light glow */}
    <div className="absolute -top-10 -right-10 w-32 h-32 bg-indigo-100/60 rounded-full blur-3xl pointer-events-none"></div>

    {(title || Icon) && (
      <div className="flex items-center space-x-3 mb-3 text-slate-500 z-10">
        {Icon && <Icon size={18} className="text-slate-400" />}
        <h3 className="font-semibold text-xs tracking-wider uppercase">{title}</h3>
      </div>
    )}
    <div className="flex-grow flex flex-col z-10">
      {children}
    </div>
  </div>
);

// Custom Dot Renderer for Recharts anomaly highlighting
const CustomDot = (props) => {
  const { cx, cy, payload } = props;
  if (payload && payload.is_anomaly) {
    return (
      <g>
        <circle cx={cx} cy={cy} r={8} fill="#dc2626" stroke="#ffffff" strokeWidth={1.5} className="animate-ping" opacity={0.5} />
        <circle cx={cx} cy={cy} r={5} fill="#dc2626" stroke="#ffffff" strokeWidth={1.5} />
      </g>
    );
  }
  return null;
};

export default function Dashboard() {
  const [currentMetrics, setCurrentMetrics] = useState(null);
  const [historyData, setHistoryData] = useState([]);
  const [isApiConnected, setIsApiConnected] = useState(false);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [modelStats, setModelStats] = useState(null);
  const [waStatus, setWaStatus] = useState(false);

  const fetchData = async (showLoadingIndicator = false) => {
    if (showLoadingIndicator) setIsRefreshing(true);
    try {
      const resCurrent = await axios.get(`${API_BASE_URL}/api/metrics/current`);
      setCurrentMetrics(resCurrent.data);
      setIsApiConnected(true);

      const resHistory = await axios.get(`${API_BASE_URL}/api/metrics/history`);
      if (Array.isArray(resHistory.data)) {
        setHistoryData(resHistory.data);
      }

      // Fetch model stats
      try {
        const resModel = await axios.get(`${API_BASE_URL}/api/model/stats`);
        setModelStats(resModel.data);
      } catch (e) { /* optional endpoint */ }

      // Fetch Telegram status
      try {
        const resTg = await axios.get(`${API_BASE_URL}/api/telegram/status`);
        setWaStatus(resTg.data.ready);
      } catch (e) { /* optional endpoint */ }
    } catch (error) {
      console.error('API connection failed. Running in standalone Mock UI preview.', error.message);
      setIsApiConnected(false);
      generateFrontendMockData();
    } finally {
      if (showLoadingIndicator) setIsRefreshing(false);
    }
  };

  const generateFrontendMockData = () => {
    const now = new Date();
    const mockIsAnomaly = Math.random() > 0.85;

    const mockCurrent = {
      is_online: true,
      is_mock: true,
      router_ip: '192.168.1.1',
      cpu_usage: mockIsAnomaly ? Math.floor(Math.random() * 15) + 80 : Math.floor(Math.random() * 25) + 12,
      uptime: '14h 28m 09s',
      rx_speed_mbps: mockIsAnomaly ? parseFloat((Math.random() * 100 + 350).toFixed(2)) : parseFloat((Math.random() * 40 + 10).toFixed(2)),
      tx_speed_mbps: mockIsAnomaly ? parseFloat((Math.random() * 30 + 80).toFixed(2)) : parseFloat((Math.random() * 10 + 2).toFixed(2)),
      is_anomaly: mockIsAnomaly,
      anomaly_reason: mockIsAnomaly ? 'Simulasi: Lonjakan Trafik / CPU Tinggi (Mock UI)' : null,
      clients: [
        { hostname: 'ASUS-ROG-Laptop', ip: '192.168.1.12', mac: 'b4:2e:99:ab:cd:ef', ssid: 'Winaila 5G' },
        { hostname: 'iPhone-15-Pro', ip: '192.168.1.15', mac: '0e:56:20:cc:4b:bd', ssid: 'Winaila 5G' },
        { hostname: 'Samsung-Smart-TV', ip: '192.168.1.18', mac: 'a4:70:d6:12:34:56', ssid: 'Winaila 2.4G' }
      ],
      client_count: 3,
      timestamp: now.getTime()
    };

    if (mockIsAnomaly && mockCurrent.anomaly_reason.includes('Trafik')) {
      mockCurrent.clients.push({ hostname: 'Unknown-Downloader-PC', ip: '192.168.1.99', mac: 'ff:ff:ff:ff:ff:ff', ssid: 'Winaila 5G' });
      mockCurrent.client_count = mockCurrent.clients.length;
    }

    setCurrentMetrics(mockCurrent);

    setHistoryData(prev => {
      const newPoint = {
        created_at: now.toISOString(),
        rx_speed_mbps: mockCurrent.rx_speed_mbps,
        tx_speed_mbps: mockCurrent.tx_speed_mbps,
        cpu_usage: mockCurrent.cpu_usage,
        is_anomaly: mockCurrent.is_anomaly,
        anomaly_reason: mockCurrent.anomaly_reason
      };
      const updated = [...prev, newPoint];
      return updated.length > 30 ? updated.slice(1) : updated;
    });
  };

  useEffect(() => {
    fetchData();
    const interval = setInterval(() => fetchData(), 5000);
    return () => clearInterval(interval);
  }, []);

  const cpuPieData = useMemo(() => {
    const usage = currentMetrics?.cpu_usage || 0;
    return [
      { name: 'Used', value: parseFloat(usage) },
      { name: 'Free', value: 100 - parseFloat(usage) }
    ];
  }, [currentMetrics?.cpu_usage]);

  const handleManualRefresh = () => {
    fetchData(true);
  };

  let statusText = 'OFFLINE';
  let statusColorClass = 'bg-red-500';
  if (currentMetrics?.is_online) {
    if (currentMetrics?.is_mock) {
      statusText = 'MOCK MODE';
      statusColorClass = 'bg-purple-500';
    } else {
      statusText = 'LIVE';
      statusColorClass = 'bg-emerald-500 animate-pulse';
    }
  }

  const formatChartTime = (timeString) => {
    if (!timeString) return '';
    const date = new Date(timeString);
    return date.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  };

  const getNetworkScore = () => {
    if (!currentMetrics) return { score: 0, grade: '?', color: 'text-slate-400', bg: 'bg-slate-50', border: 'border-slate-200' };
    let score = 100;
    if (currentMetrics.cpu_usage > 90) score -= 30;
    else if (currentMetrics.cpu_usage > 70) score -= 15;
    else if (currentMetrics.cpu_usage > 50) score -= 5;
    
    if (currentMetrics.is_anomaly) score -= 40;
    if (score < 0) score = 0;
    
    let grade = 'A', color = 'text-emerald-500', bg = 'bg-emerald-50', border = 'border-emerald-200';
    if (score >= 90) { grade = 'A'; color = 'text-emerald-500'; bg = 'bg-emerald-50'; border = 'border-emerald-200'; }
    else if (score >= 75) { grade = 'B'; color = 'text-blue-500'; bg = 'bg-blue-50'; border = 'border-blue-200'; }
    else if (score >= 50) { grade = 'C'; color = 'text-amber-500'; bg = 'bg-amber-50'; border = 'border-amber-200'; }
    else { grade = 'D'; color = 'text-red-500'; bg = 'bg-red-50'; border = 'border-red-200'; }
    return { score, grade, color, bg, border };
  };
  const netScore = getNetworkScore();

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-100 via-blue-50 to-indigo-50 text-slate-800 p-4 md:p-8 font-sans selection:bg-blue-200 relative flex flex-col justify-between">

      {/* Background Decorative Light Ornaments */}
      <div className="fixed top-[-15%] left-[-10%] w-[50%] h-[50%] rounded-full bg-blue-200/40 blur-[130px] pointer-events-none"></div>
      <div className="fixed bottom-[-15%] right-[-10%] w-[50%] h-[50%] rounded-full bg-indigo-200/40 blur-[130px] pointer-events-none"></div>

      <div className="max-w-7xl mx-auto w-full space-y-6 relative z-10 flex-grow">

        {/* HEADER SECTION */}
        <header className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white/80 backdrop-blur-xl border border-slate-200/80 p-6 rounded-3xl shadow-[0_4px_20px_rgba(100,116,139,0.10)] gap-4">
          <div className="flex items-center space-x-4">
            <div className="p-3 bg-gradient-to-br from-blue-500 to-indigo-600 rounded-2xl shadow-[0_0_15px_rgba(59,130,246,0.25)]">
              <Activity className="text-white w-6 h-6" />
            </div>
            <div>
              <h1 className="text-2xl font-bold tracking-tight bg-gradient-to-r from-slate-800 via-slate-700 to-slate-500 bg-clip-text text-transparent">
                NetVision Pro
              </h1>
              <p className="text-xs text-slate-400 mt-1 flex items-center gap-2 font-mono">
                <Server size={12} className="text-blue-500" />
                Target IP: {currentMetrics?.router_ip || '192.168.1.1'}
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-3 w-full md:w-auto justify-between md:justify-end">
            {!isApiConnected && (
              <div className="flex items-center space-x-1.5 text-amber-600 text-xs bg-amber-50 px-3 py-1.5 rounded-full border border-amber-200 font-medium">
                <AlertCircle size={13} />
                <span>Local Preview Mode</span>
              </div>
            )}
            <div className="flex items-center space-x-2 bg-slate-100 px-4 py-2 rounded-full border border-slate-200">
              <span className="relative flex h-2.5 w-2.5">
                <span className={`animate-ping absolute inline-flex h-full w-full rounded-full opacity-75 ${statusColorClass}`}></span>
                <span className={`relative inline-flex rounded-full h-2.5 w-2.5 ${statusColorClass}`}></span>
              </span>
              <span className="text-xs font-semibold uppercase tracking-wider text-slate-600">
                {statusText}
              </span>
            </div>
            <div className="flex items-center space-x-2 bg-slate-100 px-4 py-2 rounded-full border border-slate-200 text-slate-600 font-semibold text-xs">
              <Users size={14} className="text-indigo-500" />
              <span>{currentMetrics?.client_count || 0} Client</span>
            </div>
            <div className={`flex items-center space-x-2 px-3 py-2 rounded-full border font-semibold text-xs transition-colors ${waStatus ? 'bg-blue-50 border-blue-200 text-blue-600' : 'bg-slate-100 border-slate-200 text-slate-400'}`} title="Telegram Status">
              <Send size={14} className={waStatus ? "text-blue-500" : "text-slate-400"} />
              <span className="hidden md:inline">Telegram {waStatus ? 'Ready' : 'Off'}</span>
            </div>
            <button
              onClick={handleManualRefresh}
              disabled={isRefreshing}
              className="p-2 bg-slate-100 hover:bg-slate-200 active:bg-slate-300 border border-slate-200 hover:border-slate-300 rounded-full transition-colors flex items-center justify-center text-slate-500 hover:text-slate-700"
              title="Refresh Manual"
            >
              <RefreshCw size={15} className={isRefreshing ? 'animate-spin text-blue-500' : ''} />
            </button>
          </div>
        </header>

        {/* BENTO GRID LAYOUT */}
        <div className="grid grid-cols-1 md:grid-cols-4 md:grid-rows-3 gap-6 auto-rows-[minmax(180px,auto)]">

          {/* STAT 1: Download Speed */}
          <BentoCard title="Download Speed" icon={ArrowDown} className="md:col-span-1 md:row-span-1">
            <div className="flex items-end space-x-1.5 mt-2">
              <span className="text-5xl font-extrabold text-cyan-600 tabular-nums tracking-tighter">
                {currentMetrics?.rx_speed_mbps !== undefined ? currentMetrics.rx_speed_mbps.toFixed(1) : '0.0'}
              </span>
              <span className="text-slate-400 font-semibold pb-1.5 text-sm">Mbps</span>
            </div>
            <div className="mt-auto pt-4">
              <div className="h-1.5 w-full bg-slate-100 rounded-full overflow-hidden border border-slate-200">
                <div
                  className="h-full bg-cyan-500 rounded-full transition-all duration-500"
                  style={{ width: `${Math.min(((currentMetrics?.rx_speed_mbps || 0) / 100) * 100, 100)}%` }}
                ></div>
              </div>
              <div className="flex justify-between items-center text-[10px] text-slate-400 mt-2 font-mono">
                <span>Total RX:</span>
                <span>
                  {currentMetrics?.rx_bytes ? (Number(currentMetrics.rx_bytes) / (1024*1024*1024)).toFixed(2) : '0.00'} GB
                </span>
              </div>
            </div>
          </BentoCard>

          {/* STAT 2: Upload Speed */}
          <BentoCard title="Upload Speed" icon={ArrowUp} className="md:col-span-1 md:row-span-1">
            <div className="flex items-end space-x-1.5 mt-2">
              <span className="text-5xl font-extrabold text-fuchsia-600 tabular-nums tracking-tighter">
                {currentMetrics?.tx_speed_mbps !== undefined ? currentMetrics.tx_speed_mbps.toFixed(1) : '0.0'}
              </span>
              <span className="text-slate-400 font-semibold pb-1.5 text-sm">Mbps</span>
            </div>
            <div className="mt-auto pt-4">
              <div className="h-1.5 w-full bg-slate-100 rounded-full overflow-hidden border border-slate-200">
                <div
                  className="h-full bg-fuchsia-500 rounded-full transition-all duration-500"
                  style={{ width: `${Math.min(((currentMetrics?.tx_speed_mbps || 0) / 50) * 100, 100)}%` }}
                ></div>
              </div>
              <div className="flex justify-between items-center text-[10px] text-slate-400 mt-2 font-mono">
                <span>Total TX:</span>
                <span>
                  {currentMetrics?.tx_bytes ? (Number(currentMetrics.tx_bytes) / (1024*1024*1024)).toFixed(2) : '0.00'} GB
                </span>
              </div>
            </div>
          </BentoCard>

          {/* STAT 3: System Uptime */}
          <BentoCard title="System Uptime" icon={Clock} className="md:col-span-1 md:row-span-1">
            <div className="flex flex-col justify-center h-full pb-2">
              <span className="text-3xl font-bold text-slate-700 tracking-wide truncate">
                {currentMetrics?.uptime || '0h 0m 0s'}
              </span>
              <div className="flex items-center mt-3">
                <div className={`px-2.5 py-1 rounded-md text-xs font-bold tracking-wide flex items-center space-x-1 border ${
                  !currentMetrics?.ping_ms ? 'bg-slate-50 border-slate-200 text-slate-400' :
                  currentMetrics.ping_ms < 50 ? 'bg-emerald-50 border-emerald-200 text-emerald-600' :
                  currentMetrics.ping_ms < 100 ? 'bg-amber-50 border-amber-200 text-amber-600' :
                  'bg-rose-50 border-rose-200 text-rose-600'
                }`}>
                  <Activity size={12} className="mr-1" />
                  <span>Ping: {currentMetrics?.ping_ms || 0} ms</span>
                </div>
              </div>
            </div>
            <div className="text-[10px] text-slate-400 font-mono border-t border-slate-100 pt-2">
              Siklus Polling: <span className="text-blue-500">30 Detik</span>
            </div>
          </BentoCard>

          {/* STAT 4: CPU Gauge Chart */}
          <BentoCard title="CPU Load" icon={Cpu} className="md:col-span-1 md:row-span-2 flex flex-col justify-between">
            <div className="h-44 w-full mt-2 flex items-center justify-center relative">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={cpuPieData}
                    cx="50%"
                    cy="70%"
                    startAngle={180}
                    endAngle={0}
                    innerRadius={60}
                    outerRadius={80}
                    paddingAngle={0}
                    dataKey="value"
                    stroke="none"
                  >
                    {cpuPieData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={index === 0 ? COLORS.cpuUsed : COLORS.cpuFree} />
                    ))}
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div className="absolute inset-0 flex flex-col items-center justify-center top-8">
                <span className="text-3xl font-extrabold text-blue-600">
                  {currentMetrics?.cpu_usage || 0}%
                </span>
                <span className="text-[9px] text-slate-400 uppercase tracking-widest font-bold mt-0.5">CPU Load</span>
              </div>
            </div>
            <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 text-xs text-center text-slate-500">
              Prosesor: <span className={
                (currentMetrics?.cpu_usage > 85) ? 'text-red-500 font-semibold' : 
                (currentMetrics?.cpu_usage > 50) ? 'text-amber-500 font-semibold' : 
                'text-emerald-600 font-semibold'
              }>
                {(currentMetrics?.cpu_usage > 85) ? 'Kritis' : (currentMetrics?.cpu_usage > 50) ? 'Beban Sedang' : 'Normal'}
              </span>
            </div>
          </BentoCard>

          {/* MAIN GRAPH: Traffic Bandwidth */}
          <BentoCard title="Live Traffic (Bandwidth)" icon={Wifi} className="md:col-span-3 md:row-span-2">
            <div className="h-60 w-full mt-2">
              {historyData.length === 0 ? (
                <div className="w-full h-full flex flex-col items-center justify-center text-slate-400 text-sm border border-dashed border-slate-200 rounded-2xl">
                  <Activity className="w-8 h-8 text-slate-300 animate-pulse mb-2" />
                  Mengumpulkan data historis jaringan...
                </div>
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={historyData} margin={{ top: 10, right: 10, left: -25, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.15)" vertical={false} />
                    <XAxis
                      dataKey="created_at"
                      tickFormatter={formatChartTime}
                      stroke="#94a3b8"
                      fontSize={10}
                      tickLine={false}
                      axisLine={false}
                      minTickGap={40}
                      style={{ fontFamily: 'monospace' }}
                    />
                    <YAxis
                      stroke="#94a3b8"
                      fontSize={10}
                      tickLine={false}
                      axisLine={false}
                      unit=" Mb"
                      style={{ fontFamily: 'monospace' }}
                    />
                    <RechartsTooltip
                      contentStyle={{
                        backgroundColor: 'rgba(255, 255, 255, 0.97)',
                        borderColor: 'rgba(148,163,184,0.25)',
                        borderRadius: '16px',
                        boxShadow: '0 8px 24px -4px rgba(100,116,139,0.18)',
                        color: '#1e293b',
                        fontSize: '11px'
                      }}
                      labelFormatter={(label) => `Waktu: ${new Date(label).toLocaleTimeString()}`}
                    />
                    <Line
                      type="monotone"
                      dataKey="rx_speed_mbps"
                      name="Download (Mbps)"
                      stroke={COLORS.download}
                      strokeWidth={2.5}
                      dot={<CustomDot />}
                      activeDot={{ r: 6, fill: COLORS.download, stroke: '#fff', strokeWidth: 1.5 }}
                      isAnimationActive={false}
                    />
                    <Line
                      type="monotone"
                      dataKey="tx_speed_mbps"
                      name="Upload (Mbps)"
                      stroke={COLORS.upload}
                      strokeWidth={2.5}
                      dot={<CustomDot />}
                      activeDot={{ r: 6, fill: COLORS.upload, stroke: '#fff', strokeWidth: 1.5 }}
                      isAnimationActive={false}
                    />
                  </LineChart>
                </ResponsiveContainer>
              )}
            </div>

            {/* Legends */}
            <div className="flex justify-between items-center mt-3 pt-2 border-t border-slate-100 flex-wrap gap-2">
              <div className="flex space-x-6 text-xs">
                <div className="flex items-center space-x-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-cyan-500"></div>
                  <span className="text-slate-500">Download</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-fuchsia-500"></div>
                  <span className="text-slate-500">Upload</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-red-500 animate-pulse"></div>
                  <span className="text-slate-500">Titik Anomali</span>
                </div>
              </div>
              <p className="text-[10px] text-slate-400 italic">
                *Titik merah menandakan anomali (outlier statistik Z-Score &gt; {ANOMALY_THRESHOLD})
              </p>
            </div>
          </BentoCard>

          {/* NETWORK SCORE CARD */}
          <div className={`${netScore.bg} backdrop-blur-xl border ${netScore.border} rounded-3xl p-6 flex flex-col justify-between relative overflow-hidden transition-all duration-300 md:col-span-1 md:row-span-1 ${currentMetrics?.is_anomaly ? 'animate-pulse' : ''}`}>
            <div className={`absolute -top-10 -right-10 w-32 h-32 ${netScore.bg.replace('50', '100')} rounded-full blur-3xl pointer-events-none`}></div>
            <div className={`flex items-center space-x-2 mb-2 ${netScore.color} z-10`}>
              {currentMetrics?.is_anomaly ? <AlertCircle size={20} /> : <Activity size={20} />}
              <h3 className="font-bold text-xs tracking-wider uppercase">Network Score</h3>
            </div>
            
            <div className="flex-grow flex items-center justify-between py-2 z-10">
               <div className="flex flex-col">
                 <span className={`text-5xl font-extrabold tabular-nums tracking-tighter ${netScore.color}`}>{netScore.score}</span>
                 <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider mt-1">/ 100 Points</span>
               </div>
               <div className={`w-14 h-14 rounded-full flex items-center justify-center border-4 ${netScore.border} bg-white shadow-sm`}>
                 <span className={`text-2xl font-black ${netScore.color}`}>{netScore.grade}</span>
               </div>
            </div>
            
            <div className={`text-[9px] uppercase tracking-wider ${netScore.color} font-bold mt-2 font-mono border-t ${netScore.border} pt-2 truncate`}>
               {currentMetrics?.is_anomaly ? (currentMetrics.anomaly_reason || "Ada aktivitas jaringan abnormal") : "Sistem Berjalan Optimal"}
            </div>
          </div>

        </div>

        {/* MIDDLE SECTION: CPU HISTORY */}
        <div className="grid grid-cols-1 mt-6">
          <BentoCard title="CPU Load History" icon={Activity}>
            <div className="w-full h-48 pt-2">
              {historyData.length === 0 ? (
                 <div className="w-full h-full flex items-center justify-center text-slate-400 text-sm border border-dashed border-slate-200 rounded-2xl">
                    Menunggu data CPU...
                 </div>
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={historyData} margin={{ top: 10, right: 10, left: -25, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.15)" vertical={false} />
                    <XAxis
                      dataKey="created_at"
                      tickFormatter={formatChartTime}
                      stroke="#94a3b8"
                      fontSize={10}
                      tickLine={false}
                      axisLine={false}
                      minTickGap={40}
                      style={{ fontFamily: 'monospace' }}
                    />
                    <YAxis
                      stroke="#94a3b8"
                      fontSize={10}
                      tickLine={false}
                      axisLine={false}
                      unit="%"
                      domain={[0, 100]}
                      style={{ fontFamily: 'monospace' }}
                    />
                    <RechartsTooltip
                      contentStyle={{
                        backgroundColor: 'rgba(255, 255, 255, 0.97)',
                        borderColor: 'rgba(148,163,184,0.25)',
                        borderRadius: '16px',
                        boxShadow: '0 8px 24px -4px rgba(100,116,139,0.18)',
                        color: '#1e293b',
                        fontSize: '11px'
                      }}
                      labelFormatter={(label) => `Waktu: ${new Date(label).toLocaleTimeString()}`}
                    />
                    <Line
                      type="monotone"
                      dataKey="cpu_usage"
                      name="CPU Usage"
                      stroke={COLORS.cpuUsed}
                      strokeWidth={2.5}
                      dot={false}
                      activeDot={{ r: 6, fill: COLORS.cpuUsed, stroke: '#fff', strokeWidth: 1.5 }}
                      isAnimationActive={false}
                    />
                  </LineChart>
                </ResponsiveContainer>
              )}
            </div>
          </BentoCard>
        </div>

        {/* BOTTOM DETAILS */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mt-6">
          <BentoCard title="Metode Deteksi" icon={Info} className="lg:col-span-1">
            <div className="text-xs text-slate-500 space-y-2 leading-relaxed">
              <p>
                Sistem mendeteksi pencilan jaringan menggunakan algoritma <strong className="text-slate-700">Z-Score Multi-Dimensi</strong>.
              </p>
              {modelStats?.is_trained ? (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1.5 font-mono text-[10px]">
                  <div className="flex justify-between"><span className="text-slate-400">CPU μ/σ:</span><span className="text-slate-700">{modelStats.cpu.mean.toFixed(1)}% / {modelStats.cpu.stdDev.toFixed(1)}%</span></div>
                  <div className="flex justify-between"><span className="text-slate-400">Rx μ/σ:</span><span className="text-cyan-600">{modelStats.rx.mean.toFixed(2)} / {modelStats.rx.stdDev.toFixed(2)} Mbps</span></div>
                  <div className="flex justify-between"><span className="text-slate-400">Tx μ/σ:</span><span className="text-fuchsia-600">{modelStats.tx.mean.toFixed(2)} / {modelStats.tx.stdDev.toFixed(2)} Mbps</span></div>
                  <div className="border-t border-slate-200 pt-1.5 mt-1.5">
                    <div className="flex justify-between"><span className="text-slate-400">Z-CPU:</span><span className={modelStats.current_z_scores.cpu > ANOMALY_THRESHOLD ? 'text-red-500 font-bold' : 'text-slate-700'}>{modelStats.current_z_scores.cpu}</span></div>
                    <div className="flex justify-between"><span className="text-slate-400">Z-Rx:</span><span className={modelStats.current_z_scores.rx > ANOMALY_THRESHOLD ? 'text-red-500 font-bold' : 'text-slate-700'}>{modelStats.current_z_scores.rx}</span></div>
                    <div className="flex justify-between"><span className="text-slate-400">Z-Tx:</span><span className={modelStats.current_z_scores.tx > ANOMALY_THRESHOLD ? 'text-red-500 font-bold' : 'text-slate-700'}>{modelStats.current_z_scores.tx}</span></div>
                  </div>
                  <p className="text-[9px] text-slate-400 pt-1">Threshold: Z &gt; {ANOMALY_THRESHOLD} = Anomali</p>
                </div>
              ) : (
                <p>
                  Model dilatih setiap 5 menit mengambil 100 riwayat log Supabase untuk menghitung standar deviasi dan rata-rata.
                </p>
              )}
              <p>
                Jika metrik baru menyimpang melampaui <strong className="text-blue-600">{ANOMALY_THRESHOLD}x standar deviasi</strong>, status anomali dipicu secara otomatis.
              </p>
            </div>
          </BentoCard>

          <BentoCard title={`Client Terkoneksi (${currentMetrics?.client_count || 0})`} icon={Users} className="lg:col-span-1">
            <div className="overflow-y-auto max-h-[280px] pr-1 space-y-2.5">
              {!currentMetrics?.clients || currentMetrics.clients.length === 0 ? (
                <p className="text-xs text-slate-400 italic text-center py-12">
                  Tidak ada client WiFi terdeteksi.
                </p>
              ) : (
                currentMetrics.clients.map((client, idx) => (
                  <div key={idx} className="flex items-center gap-3 bg-slate-50 p-3 rounded-2xl border border-slate-100">
                    {/* Avatar circle */}
                    <div className="w-9 h-9 rounded-full bg-gradient-to-br from-indigo-100 to-blue-50 border border-indigo-200 flex items-center justify-center shrink-0">
                      <Wifi size={14} className="text-indigo-500" />
                    </div>
                    {/* Info */}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-semibold text-xs text-slate-700 truncate">{client.hostname || 'Unknown'}</span>
                        <span className="text-[9px] bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded-full font-medium shrink-0 border border-indigo-100">
                          {client.ssid || 'WiFi'}
                        </span>
                      </div>
                      <div className="flex items-center justify-between mt-1">
                        <span className="text-[10px] text-slate-400 font-mono truncate">{client.ip} • {client.mac}</span>
                        {client.active_since && (
                          <span className="text-[9px] text-slate-400 shrink-0 ml-2">
                            Since {new Date(client.active_since).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}
                          </span>
                        )}
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </BentoCard>

          <BentoCard title="Log Database Historis (Supabase)" icon={History} className="lg:col-span-1">
            <div className="overflow-y-auto max-h-[280px] pr-1">
              <div className="space-y-2">
                {historyData.length === 0 ? (
                  <p className="text-xs text-slate-400 italic text-center py-12">
                    Belum ada log jaringan di database.
                  </p>
                ) : (
                  [...historyData].reverse().slice(0, 12).map((log, idx) => {
                    const date = new Date(log.created_at);
                    const time = date.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
                    const day = date.toLocaleDateString('id-ID', { day: 'numeric', month: 'short' });
                    return (
                      <div key={log.id || idx} className={`p-2.5 rounded-xl border text-[11px] ${log.is_anomaly ? 'bg-red-50/50 border-red-100' : 'bg-slate-50/50 border-slate-100'}`}>
                        <div className="flex items-center justify-between mb-1">
                          <span className="text-slate-400 font-mono text-[10px]">{time} {day}</span>
                          {log.is_anomaly ? (
                            <span className="text-[9px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded font-bold">ANOMALI</span>
                          ) : (
                            <span className="text-[9px] bg-emerald-100 text-emerald-600 px-1.5 py-0.5 rounded font-bold">NORMAL</span>
                          )}
                        </div>
                        <div className="flex items-center gap-3 font-mono">
                          <span className="text-slate-500">CPU <strong className="text-slate-700">{log.cpu_usage}%</strong></span>
                          <span className="text-cyan-600">↓{(log.rx_speed_mbps !== undefined && log.rx_speed_mbps !== null) ? log.rx_speed_mbps.toFixed(1) : '0.0'}</span>
                          <span className="text-fuchsia-600">↑{(log.tx_speed_mbps !== undefined && log.tx_speed_mbps !== null) ? log.tx_speed_mbps.toFixed(1) : '0.0'}</span>
                          <span className="text-slate-400 text-[10px]">Mbps</span>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          </BentoCard>
        </div>

      </div>

      <footer className="text-center text-[10px] text-slate-400 mt-8 pt-4 border-t border-slate-200 max-w-7xl mx-auto w-full">
        NetVision Pro • Monitoring & Deteksi Anomali Jaringan Pintar • © {new Date().getFullYear()}
      </footer>
    </div>
  );
}
