import React, { useState, useEffect, useMemo } from 'react';
import { 
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { Activity, ArrowDown, ArrowUp, Cpu, Clock, Server, Wifi, AlertCircle } from 'lucide-react';

const API_BASE_URL = 'http://localhost:5000/api';

// Warna Neon untuk Chart
const COLORS = {
  download: '#06b6d4', // Cyan
  upload: '#d946ef',   // Fuchsia
  cpuUsed: '#10b981',  // Emerald
  cpuFree: '#1f2937'   // Gray 800
};

// Komponen Card Bento dengan Glassmorphism
const BentoCard = ({ children, className = '', title, icon: Icon }) => (
  <div className={`bg-gray-900/40 backdrop-blur-xl border border-white/10 rounded-3xl p-6 shadow-[0_8px_32px_0_rgba(0,0,0,0.3)] flex flex-col relative overflow-hidden transition-all duration-300 hover:border-white/20 hover:bg-gray-800/40 ${className}`}>
    {/* Gradasi halus di latar belakang kartu */}
    <div className="absolute -top-10 -right-10 w-32 h-32 bg-white/5 rounded-full blur-3xl pointer-events-none"></div>
    
    {(title || Icon) && (
      <div className="flex items-center space-x-3 mb-4 text-gray-400 z-10">
        {Icon && <Icon size={20} className="text-gray-300" />}
        <h3 className="font-medium text-sm tracking-wider uppercase">{title}</h3>
      </div>
    )}
    <div className="flex-grow flex flex-col z-10">
      {children}
    </div>
  </div>
);

export default function NetworkDashboard() {
  const [currentMetrics, setCurrentMetrics] = useState(null);
  const [historyData, setHistoryData] = useState([]);
  const [isApiConnected, setIsApiConnected] = useState(false);

  useEffect(() => {
    // Fungsi untuk menghasilkan data dummy jika Backend tidak tersedia (berguna untuk UI Preview)
    const generateMockData = () => {
      const now = new Date();
      return {
        status: 'online',
        router_ip: '192.168.1.1',
        cpu_usage: Math.floor(Math.random() * 40) + 10, // 10% - 50%
        uptime: '12h 45m 22s',
        rx_mbps: parseFloat((Math.random() * 150 + 50).toFixed(2)), // 50 - 200 Mbps
        tx_mbps: parseFloat((Math.random() * 80 + 20).toFixed(2)),  // 20 - 100 Mbps
        timestamp: now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      };
    };

    // Inisialisasi Mock History
    const initMockHistory = () => {
      const mockHistory = [];
      let time = new Date(Date.now() - 30 * 60000); // 30 mins ago
      for (let i = 0; i < 30; i++) {
        mockHistory.push({
          time: time.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
          rx_mbps: Math.random() * 150 + 50,
          tx_mbps: Math.random() * 80 + 20,
        });
        time = new Date(time.getTime() + 60000);
      }
      setHistoryData(mockHistory);
    };

    const fetchData = async () => {
      try {
        const res = await fetch(`${API_BASE_URL}/metrics/current`);
        if (!res.ok) throw new Error('API Error');
        const data = await res.json();
        
        setCurrentMetrics(data);
        setIsApiConnected(true);
        
        // Update history chart
        setHistoryData(prev => {
          const newPoint = {
            time: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second:'2-digit' }),
            rx_mbps: data.rx_mbps,
            tx_mbps: data.tx_mbps
          };
          const newHistory = [...prev, newPoint];
          return newHistory.length > 50 ? newHistory.slice(1) : newHistory;
        });

      } catch (error) {
        // Fallback ke Mock Data jika API tidak merespon
        setIsApiConnected(false);
        const mock = generateMockData();
        setCurrentMetrics(mock);
        
        setHistoryData(prev => {
          if (prev.length === 0) return prev;
          const newPoint = {
            time: mock.timestamp,
            rx_mbps: mock.rx_mbps,
            tx_mbps: mock.tx_mbps
          };
          const newHistory = [...prev, newPoint];
          return newHistory.length > 30 ? newHistory.slice(1) : newHistory;
        });
      }
    };

    // Panggil init mock jika pertama kali render
    initMockHistory();
    fetchData(); // Fetch langsung
    const interval = setInterval(fetchData, 5000); // Polling tiap 5 detik
    return () => clearInterval(interval);
  }, []);

  const cpuData = useMemo(() => {
    const usage = currentMetrics?.cpu_usage || 0;
    return [
      { name: 'Used', value: usage },
      { name: 'Free', value: 100 - usage }
    ];
  }, [currentMetrics?.cpu_usage]);

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-slate-200 p-4 md:p-8 font-sans selection:bg-cyan-900 overflow-x-hidden relative">
      
      {/* Background Ornaments */}
      <div className="fixed top-[-20%] left-[-10%] w-[50%] h-[50%] rounded-full bg-cyan-900/20 blur-[120px] pointer-events-none"></div>
      <div className="fixed bottom-[-20%] right-[-10%] w-[50%] h-[50%] rounded-full bg-fuchsia-900/10 blur-[120px] pointer-events-none"></div>

      <div className="max-w-7xl mx-auto space-y-6 relative z-10">
        
        {/* Header Section */}
        <header className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white/5 backdrop-blur-md border border-white/10 p-5 rounded-3xl shadow-lg">
          <div className="flex items-center space-x-4">
            <div className="p-3 bg-gradient-to-br from-cyan-500 to-blue-600 rounded-2xl shadow-[0_0_15px_rgba(6,182,212,0.5)]">
              <Activity className="text-white w-6 h-6" />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-white tracking-tight">NetVision Pro</h1>
              <p className="text-sm text-gray-400 mt-1 flex items-center gap-2">
                <Server size={14} /> 
                {currentMetrics?.router_ip || 'Loading...'}
              </p>
            </div>
          </div>

          <div className="mt-4 md:mt-0 flex items-center space-x-4">
            {!isApiConnected && (
              <div className="flex items-center space-x-2 text-yellow-500 text-xs bg-yellow-500/10 px-3 py-1.5 rounded-full border border-yellow-500/20">
                <AlertCircle size={14} />
                <span>Mock Mode (API Backend Offline)</span>
              </div>
            )}
            <div className="flex items-center space-x-2 bg-gray-900/50 px-4 py-2 rounded-full border border-gray-700">
              <span className="relative flex h-3 w-3">
                <span className={`animate-ping absolute inline-flex h-full w-full rounded-full opacity-75 ${currentMetrics?.status === 'online' ? 'bg-green-400' : 'bg-red-400'}`}></span>
                <span className={`relative inline-flex rounded-full h-3 w-3 ${currentMetrics?.status === 'online' ? 'bg-green-500' : 'bg-red-500'}`}></span>
              </span>
              <span className="text-sm font-medium uppercase tracking-wider text-gray-300">
                {currentMetrics?.status || 'Connecting'}
              </span>
            </div>
          </div>
        </header>

        {/* Bento Grid Layout */}
        <div className="grid grid-cols-1 md:grid-cols-4 md:grid-rows-3 gap-6 auto-rows-[minmax(180px,auto)]">
          
          {/* STAT: Download Speed */}
          <BentoCard title="Download Speed" icon={ArrowDown} className="md:col-span-1 md:row-span-1">
            <div className="flex items-end space-x-2 mt-2">
              <span className="text-5xl font-bold text-cyan-400 tabular-nums tracking-tighter">
                {currentMetrics?.rx_mbps?.toFixed(1) || '0.0'}
              </span>
              <span className="text-gray-400 font-medium pb-1">Mbps</span>
            </div>
            <div className="mt-auto pt-4 flex items-center text-xs text-cyan-500/80">
              <div className="h-1 w-full bg-gray-800 rounded-full overflow-hidden">
                <div className="h-full bg-cyan-400 rounded-full" style={{ width: `${Math.min(((currentMetrics?.rx_mbps||0)/1000)*100, 100)}%` }}></div>
              </div>
            </div>
          </BentoCard>

          {/* STAT: Upload Speed */}
          <BentoCard title="Upload Speed" icon={ArrowUp} className="md:col-span-1 md:row-span-1">
            <div className="flex items-end space-x-2 mt-2">
              <span className="text-5xl font-bold text-fuchsia-400 tabular-nums tracking-tighter">
                {currentMetrics?.tx_mbps?.toFixed(1) || '0.0'}
              </span>
              <span className="text-gray-400 font-medium pb-1">Mbps</span>
            </div>
            <div className="mt-auto pt-4 flex items-center text-xs text-fuchsia-500/80">
               <div className="h-1 w-full bg-gray-800 rounded-full overflow-hidden">
                <div className="h-full bg-fuchsia-400 rounded-full" style={{ width: `${Math.min(((currentMetrics?.tx_mbps||0)/1000)*100, 100)}%` }}></div>
              </div>
            </div>
          </BentoCard>

          {/* STAT: Uptime */}
          <BentoCard title="System Uptime" icon={Clock} className="md:col-span-1 md:row-span-1">
            <div className="flex items-center h-full">
              <span className="text-3xl font-semibold text-white tracking-wide">
                {currentMetrics?.uptime || '--:--:--'}
              </span>
            </div>
          </BentoCard>

          {/* CPU GAUGE CHART */}
          <BentoCard title="CPU Load" icon={Cpu} className="md:col-span-1 md:row-span-2 flex flex-col justify-between">
            <div className="h-48 w-full mt-4 flex items-center justify-center relative">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={cpuData}
                    cx="50%"
                    cy="70%"
                    startAngle={180}
                    endAngle={0}
                    innerRadius={65}
                    outerRadius={85}
                    paddingAngle={2}
                    dataKey="value"
                    stroke="none"
                  >
                    {cpuData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={index === 0 ? COLORS.cpuUsed : COLORS.cpuFree} />
                    ))}
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div className="absolute inset-0 flex flex-col items-center justify-center top-10">
                <span className="text-4xl font-bold text-emerald-400 drop-shadow-[0_0_8px_rgba(16,185,129,0.5)]">
                  {currentMetrics?.cpu_usage || 0}%
                </span>
                <span className="text-xs text-gray-400 mt-1 uppercase tracking-widest">Load</span>
              </div>
            </div>
            <div className="text-sm text-center text-gray-400 mt-2">
              Status: <span className={(currentMetrics?.cpu_usage > 80) ? 'text-red-400' : 'text-emerald-400'}>
                {(currentMetrics?.cpu_usage > 80) ? 'Critical' : 'Normal'}
              </span>
            </div>
          </BentoCard>

          {/* MAIN CHART: Traffic History */}
          <BentoCard title="Live Traffic (Bandwidth)" icon={Wifi} className="md:col-span-3 md:row-span-2">
            <div className="h-64 w-full mt-2">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={historyData} margin={{ top: 5, right: 0, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#334155" vertical={false} />
                  <XAxis 
                    dataKey="time" 
                    stroke="#64748b" 
                    fontSize={12} 
                    tickLine={false}
                    axisLine={false}
                    minTickGap={30}
                  />
                  <YAxis 
                    stroke="#64748b" 
                    fontSize={12} 
                    tickLine={false}
                    axisLine={false}
                    tickFormatter={(value) => `${value}`}
                  />
                  <RechartsTooltip 
                    contentStyle={{ backgroundColor: '#1e293b', border: '1px solid rgba(255,255,255,0.1)', borderRadius: '12px', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.5)' }}
                    itemStyle={{ color: '#e2e8f0' }}
                    labelStyle={{ color: '#94a3b8', marginBottom: '4px' }}
                  />
                  <Line 
                    type="monotone" 
                    dataKey="rx_mbps" 
                    name="Download (Mbps)" 
                    stroke={COLORS.download} 
                    strokeWidth={3} 
                    dot={false}
                    activeDot={{ r: 6, fill: COLORS.download, stroke: '#fff', strokeWidth: 2 }}
                    isAnimationActive={false} // Disable animation for smoother real-time feel
                  />
                  <Line 
                    type="monotone" 
                    dataKey="tx_mbps" 
                    name="Upload (Mbps)" 
                    stroke={COLORS.upload} 
                    strokeWidth={3} 
                    dot={false}
                    activeDot={{ r: 6, fill: COLORS.upload, stroke: '#fff', strokeWidth: 2 }}
                    isAnimationActive={false}
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
            
            {/* Legend Kustom */}
            <div className="flex justify-center space-x-6 mt-4">
              <div className="flex items-center space-x-2">
                <div className="w-3 h-3 rounded-full bg-cyan-400 shadow-[0_0_10px_rgba(6,182,212,0.8)]"></div>
                <span className="text-sm text-gray-300">Download</span>
              </div>
              <div className="flex items-center space-x-2">
                <div className="w-3 h-3 rounded-full bg-fuchsia-400 shadow-[0_0_10px_rgba(217,70,239,0.8)]"></div>
                <span className="text-sm text-gray-300">Upload</span>
              </div>
            </div>
          </BentoCard>

        </div>
      </div>
    </div>
  );
}