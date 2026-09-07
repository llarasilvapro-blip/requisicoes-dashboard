import React, { useState } from 'react';
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  BarChart, Bar, PieChart, Pie, Cell 
} from 'recharts';

// --- DADOS DO DASHBOARD ---
const evolucaoData = [
  { mes: '2020-05', solicitacoes: 18, modificacoes: 5 },
  { mes: '2020-06', solicitacoes: 65, modificacoes: 12 },
  { mes: '2020-07', solicitacoes: 125, modificacoes: 30 },
  { mes: '2020-08', solicitacoes: 235, modificacoes: 310 },
  { mes: '2020-09', solicitacoes: 40, modificacoes: 45 },
];

const agingBarData = [
  { faixa: '0 - 15 dias', itens: 183, rcs: 104 },
  { faixa: '16 - 30 dias', itens: 84, rcs: 54 },
  { faixa: '31 - 60 dias', itens: 79, rcs: 58 },
  { faixa: 'Acima de 60 dias', itens: 44, rcs: 18 },
];

const remessaData = [
  { name: 'Remessa no prazo', value: 82, color: '#00E599' },
  { name: 'Remessa atrasada', value: 12, color: '#F59E0B' },
  { name: 'Sem data', value: 6, color: '#A855F7' },
];

export default function App() {
  const [filtroCentro, setFiltroCentro] = useState('Todos');
  const [filtroComprador, setFiltroComprador] = useState('Todos');
  const [filtroGC, setFiltroGC] = useState('Todos');
  const [filtroAging, setFiltroAging] = useState('Todos');

  return (
    <div className="min-h-screen bg-[#0A0E17] text-[#E2E8F0] p-6 font-sans">
      
      {/* CABEÇALHO SUPERIOR */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4">
        <div>
          <p className="text-[11px] font-bold tracking-wider text-[#00E599] uppercase mb-0.5">
            PROCUREMENT · REQUISIÇÕES DE COMPRA
          </p>
          <h1 className="text-2xl font-bold text-white tracking-tight">
            Dashboard de Requisições
          </h1>
          <p className="text-xs text-slate-400 mt-0.5">
            Base padrão (planilha enviada)
          </p>
        </div>

        <div className="flex items-center gap-3">
          <button className="px-4 py-2 bg-[#171E2E] hover:bg-[#232D42] text-xs font-medium text-slate-300 rounded-lg border border-slate-800 transition">
            Divergências
          </button>
          <button className="px-4 py-2 bg-[#00E599] hover:bg-[#00C282] text-xs font-bold text-[#0A0E17] rounded-lg transition shadow-lg shadow-[#00E599]/10">
            Atualizar base (.xlsx)
          </button>
        </div>
      </div>

      {/* BARRA DE FILTROS */}
      <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80 mb-6 flex flex-col md:flex-row justify-between items-center gap-4">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 w-full md:w-auto flex-1 max-w-4xl">
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Centro</label>
            <select 
              value={filtroCentro} onChange={(e) => setFiltroCentro(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              <option value="Todos">Todos</option>
            </select>
          </div>

          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Comprador</label>
            <select 
              value={filtroComprador} onChange={(e) => setFiltroComprador(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              <option value="Todos">Todos</option>
            </select>
          </div>

          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Grupo de compras</label>
            <select 
              value={filtroGC} onChange={(e) => setFiltroGC(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              <option value="Todos">Todos</option>
            </select>
          </div>

          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Faixa de aging</label>
            <select 
              value={filtroAging} onChange={(e) => setFiltroAging(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              <option value="Todos">Todos</option>
            </select>
          </div>
        </div>

        <div className="text-right">
          <span className="text-xs font-semibold text-slate-400">390 de 390 itens</span>
        </div>
      </div>

      {/* CARDS DE KPIS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">VOLUME DE REQUIÇÕES</p>
          <h3 className="text-2xl font-black text-[#00E599] mt-1">232 RCs</h3>
          <p className="text-xs text-slate-400 mt-1">390 itens solicitados</p>
        </div>

        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">LEAD TIME MÉDIO (DIAS RC)</p>
          <h3 className="text-2xl font-black text-white mt-1">24,3 dias</h3>
          <p className="text-xs text-slate-400 mt-1">Pico de 85 dias</p>
        </div>

        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">PENDÊNCIAS CRÍTICAS</p>
          <h3 className="text-2xl font-black text-white mt-1">32,8%</h3>
          <p className="text-xs text-slate-400 mt-1">76 RCs &gt; 30 dias · 18 &gt; 60 dias</p>
        </div>

        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">COBERTURA OPERACIONAL</p>
          <h3 className="text-2xl font-black text-white mt-1">40 centros</h3>
          <p className="text-xs text-slate-400 mt-1">67.719 unidades solicitadas</p>
        </div>
      </div>

      {/* EVOLUÇÃO TEMPORAL */}
      <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80 mb-6">
        <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">EVOLUÇÃO TEMPORAL</h2>
        <p className="text-[11px] text-slate-400 mb-4">Solicitações abertas x modificações registradas por mês</p>
        
        <div className="h-64">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={evolucaoData}>
              <defs>
                <linearGradient id="colorSolicitacoes" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#00E599" stopOpacity={0.2}/>
                  <stop offset="95%" stopColor="#00E599" stopOpacity={0}/>
                </linearGradient>
                <linearGradient id="colorModificacoes" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#F59E0B" stopOpacity={0.2}/>
                  <stop offset="95%" stopColor="#F59E0B" stopOpacity={0}/>
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" vertical={false} />
              <XAxis dataKey="mes" stroke="#64748B" fontSize={10} />
              <YAxis stroke="#64748B" fontSize={10} />
              <Tooltip contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }} />
              <Area type="monotone" dataKey="solicitacoes" stroke="#00E599" strokeWidth={2} fillOpacity={1} fill="url(#colorSolicitacoes)" name="Solicitações" />
              <Area type="monotone" dataKey="modificacoes" stroke="#F59E0B" strokeWidth={2} fillOpacity={1} fill="url(#colorModificacoes)" name="Modificações" />
            </AreaChart>
          </ResponsiveContainer>
        </div>
        <div className="flex justify-center gap-6 mt-2 text-xs">
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-[#00E599]"></span>
            <span className="text-slate-400">Solicitações</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-[#F59E0B]"></span>
            <span className="text-slate-400">Modificações</span>
          </div>
        </div>
      </div>

      {/* AGING E CUMPRIMENTO DE DATA */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
        {/* DISTRIBUIÇÃO POR FAIXA DE AGING */}
        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">DISTRIBUIÇÃO POR FAIXA DE AGING</h2>
          <p className="text-[11px] text-slate-400 mb-4">Itens e RCs por tempo em aberto</p>
          
          <div className="h-60">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={agingBarData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" vertical={false} />
                <XAxis dataKey="faixa" stroke="#64748B" fontSize={10} />
                <YAxis stroke="#64748B" fontSize={10} />
                <Tooltip contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }} />
                <Bar dataKey="itens" fill="#00E599" radius={[3, 3, 0, 0]} name="Itens" />
                <Bar dataKey="rcs" fill="#F59E0B" radius={[3, 3, 0, 0]} name="RCs" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="flex justify-center gap-6 mt-2 text-xs">
            <div className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded bg-[#00E599]"></span>
              <span className="text-slate-400">Itens</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded bg-[#F59E0B]"></span>
              <span className="text-slate-400">RCs</span>
            </div>
          </div>
        </div>

        {/* CUMPRIMENTO DE DATA DE REMESSA */}
        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">CUMPRIMENTO DE DATA DE REMESSA</h2>
          <p className="text-[11px] text-slate-400 mb-2">Comparação entre Modificado em e Remessas (dia útil)</p>
          
          <div className="h-56">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie 
                  data={remessaData} 
                  innerRadius={65} 
                  outerRadius={85} 
                  paddingAngle={4} 
                  dataKey="value"
                >
                  {remessaData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
                <Tooltip contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="flex flex-wrap justify-center gap-4 mt-1 text-xs">
            {remessaData.map((item, idx) => (
              <div key={idx} className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: item.color }}></span>
                <span className="text-slate-400">{item.name}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* NOVA MATRIZ DE PRIORIZAÇÃO OPERACIONAL */}
      <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80 mb-6">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-5 gap-2">
          <div>
            <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-[#00E599]"></span>
              MATRIZ DE PRIORIZAÇÃO OPERACIONAL POR PLANTA
            </h2>
            <p className="text-[11px] text-slate-400 mt-0.5">
              Cruzamento de Impacto (Volume de Itens) x Severidade de Atraso (Aging em Dias)
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 text-[11px]">
            <span className="px-2.5 py-1 rounded bg-rose-500/10 text-rose-400 border border-rose-500/20 font-semibold">
              🔴 4 Críticas
            </span>
            <span className="px-2.5 py-1 rounded bg-amber-500/10 text-amber-400 border border-amber-500/20 font-semibold">
              🟠 4 Gargalos
            </span>
            <span className="px-2.5 py-1 rounded bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 font-semibold">
              🟢 4 Saudáveis
            </span>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          
          {/* QUADRANTE 1: ALTO IMPACTO + ALTO AGING */}
          <div className="bg-[#171E2E] p-4 rounded-xl border border-rose-500/30">
            <div className="flex justify-between items-center mb-3 pb-2 border-b border-rose-500/20">
              <h3 className="text-xs font-bold text-rose-400 uppercase tracking-wide">
                1. Alto Impacto & Atraso Crítico
              </h3>
              <span className="text-[10px] bg-rose-500/20 text-rose-300 px-1.5 py-0.5 rounded font-mono">
                Foco Imediato
              </span>
            </div>
            <p className="text-[10px] text-slate-400 mb-3">Plantas com grande volume represado E tempo médio alto.</p>
            
            <div className="space-y-2.5">
              {[
                { nome: 'Indústria - Pará de Minas', itens: 46, dias: 27.7 },
                { nome: 'Fábrica - Uberlândia', itens: 43, dias: 35.3 },
                { nome: 'Indústria - Araras', itens: 41, dias: 16.8 },
                { nome: 'Indústria - Sete Lagoas', itens: 29, dias: 37.0 },
              ].map((item, idx) => (
                <div key={idx} className="bg-[#101522] p-2.5 rounded-lg border border-slate-800 flex justify-between items-center">
                  <div>
                    <p className="text-xs font-semibold text-white">{item.nome}</p>
                    <p className="text-[10px] text-slate-400">{item.itens} itens aguardando</p>
                  </div>
                  <div className="text-right">
                    <span className="text-xs font-bold text-rose-400 font-mono">{item.dias} dias</span>
                    <div className="w-16 bg-slate-800 h-1 rounded-full mt-1 overflow-hidden">
                      <div className="bg-rose-500 h-full" style={{ width: `${Math.min((item.dias/40)*100, 100)}%` }}></div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* QUADRANTE 2: ALTO AGING PONTUAL */}
          <div className="bg-[#171E2E] p-4 rounded-xl border border-amber-500/30">
            <div className="flex justify-between items-center mb-3 pb-2 border-b border-amber-500/20">
              <h3 className="text-xs font-bold text-amber-400 uppercase tracking-wide">
                2. Risco de Trava / Aging Elevado
              </h3>
              <span className="text-[10px] bg-amber-500/20 text-amber-300 px-1.5 py-0.5 rounded font-mono">
                Gargalos
              </span>
            </div>
            <p className="text-[10px] text-slate-400 mb-3">Médio/baixo volume, mas com dias extremamente elevados.</p>

            <div className="space-y-2.5">
              {[
                { nome: 'Indústria - Londrina', itens: 19, dias: 46.4 },
                { nome: 'Indústria - Garanhuns', itens: 8, dias: 49.3 },
                { nome: 'Indústria - Ijuí', itens: 10, dias: 28.6 },
                { nome: 'Posto de Coleta - Marau', itens: 2, dias: 31.5 },
              ].map((item, idx) => (
                <div key={idx} className="bg-[#101522] p-2.5 rounded-lg border border-slate-800 flex justify-between items-center">
                  <div>
                    <p className="text-xs font-semibold text-white">{item.nome}</p>
                    <p className="text-[10px] text-slate-400">{item.itens} itens aguardando</p>
                  </div>
                  <div className="text-right">
                    <span className="text-xs font-bold text-amber-400 font-mono">{item.dias} dias</span>
                    <div className="w-16 bg-slate-800 h-1 rounded-full mt-1 overflow-hidden">
                      <div className="bg-amber-500 h-full" style={{ width: `${Math.min((item.dias/50)*100, 100)}%` }}></div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* QUADRANTE 3: ALTO VOLUME EM FLUXO NORMAL */}
          <div className="bg-[#171E2E] p-4 rounded-xl border border-emerald-500/30">
            <div className="flex justify-between items-center mb-3 pb-2 border-b border-emerald-500/20">
              <h3 className="text-xs font-bold text-[#00E599] uppercase tracking-wide">
                3. Alto Volume em Giro Rápido
              </h3>
              <span className="text-[10px] bg-[#00E599]/20 text-[#00E599] px-1.5 py-0.5 rounded font-mono">
                Saudável
              </span>
            </div>
            <p className="text-[10px] text-slate-400 mb-3">Grande giro de itens, mas dentro da janela aceitável de SLA.</p>

            <div className="space-y-2.5">
              {[
                { nome: 'Indústria - Carambeí', itens: 30, dias: 17.7 },
                { nome: 'Indústria - Três de Maio', itens: 20, dias: 18.6 },
                { nome: 'Indústria - Barra Mansa', itens: 19, dias: 16.3 },
                { nome: 'Indústria - Pouso Alto', itens: 19, dias: 13.7 },
              ].map((item, idx) => (
                <div key={idx} className="bg-[#101522] p-2.5 rounded-lg border border-slate-800 flex justify-between items-center">
                  <div>
                    <p className="text-xs font-semibold text-white">{item.nome}</p>
                    <p className="text-[10px] text-slate-400">{item.itens} itens aguardando</p>
                  </div>
                  <div className="text-right">
                    <span className="text-xs font-bold text-[#00E599] font-mono">{item.dias} dias</span>
                    <div className="w-16 bg-slate-800 h-1 rounded-full mt-1 overflow-hidden">
                      <div className="bg-[#00E599] h-full" style={{ width: `${Math.min((item.dias/20)*100, 100)}%` }}></div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

        </div>
      </div>

    </div>
  );
}