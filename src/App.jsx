import React, { useState } from 'react';
import * as XLSX from 'xlsx';
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
  { faixa: '0 - 15 dias', itens: 183, rcs: 104, status: 'Saudável' },
  { faixa: '16 - 30 dias', itens: 84, rcs: 54, status: 'Atenção' },
  { faixa: '31 - 60 dias', itens: 79, rcs: 58, status: 'Atrasado' },
  { faixa: 'Acima de 60 dias', itens: 44, rcs: 18, status: 'Crítico' },
];

const remessaData = [
  { name: 'Remessa no prazo', value: 82, color: '#00E599' },
  { name: 'Remessa atrasada', value: 12, color: '#F59E0B' },
  { name: 'Sem data', value: 6, color: '#A855F7' },
];

const matrizPlantas = [
  { Planta: 'Indústria - Pará de Minas', Itens: 46, AgingMedioDias: 27.7, Classificacao: '🔴 1. Alto Impacto & Atraso Crítico' },
  { Planta: 'Fábrica - Uberlândia', Itens: 43, AgingMedioDias: 35.3, Classificacao: '🔴 1. Alto Impacto & Atraso Crítico' },
  { Planta: 'Indústria - Araras', Itens: 41, AgingMedioDias: 16.8, Classificacao: '🔴 1. Alto Impacto & Atraso Crítico' },
  { Planta: 'Indústria - Sete Lagoas', Itens: 29, AgingMedioDias: 37.0, Classificacao: '🔴 1. Alto Impacto & Atraso Crítico' },
  { Planta: 'Indústria - Londrina', Itens: 19, AgingMedioDias: 46.4, Classificacao: '🟠 2. Risco de Trava / Gargalo' },
  { Planta: 'Indústria - Garanhuns', Itens: 8, AgingMedioDias: 49.3, Classificacao: '🟠 2. Risco de Trava / Gargalo' },
  { Planta: 'Indústria - Ijuí', Itens: 10, AgingMedioDias: 28.6, Classificacao: '🟠 2. Risco de Trava / Gargalo' },
  { Planta: 'Posto de Coleta - Marau', Itens: 2, AgingMedioDias: 31.5, Classificacao: '🟠 2. Risco de Trava / Gargalo' },
  { Planta: 'Indústria - Carambeí', Itens: 30, AgingMedioDias: 17.7, Classificacao: '🟢 3. Alto Volume em Giro Rápido' },
  { Planta: 'Indústria - Três de Maio', Itens: 20, AgingMedioDias: 18.6, Classificacao: '🟢 3. Alto Volume em Giro Rápido' },
  { Planta: 'Indústria - Barra Mansa', Itens: 19, AgingMedioDias: 16.3, Classificacao: '🟢 3. Alto Volume em Giro Rápido' },
  { Planta: 'Indústria - Pouso Alto', Itens: 19, AgingMedioDias: 13.7, Classificacao: '🟢 3. Alto Volume em Giro Rápido' },
];

export default function App() {
  const [filtroCentro, setFiltroCentro] = useState('Todos');
  const [filtroComprador, setFiltroComprador] = useState('Todos');
  const [filtroGC, setFiltroGC] = useState('Todos');
  const [filtroAging, setFiltroAging] = useState('Todos');
  const [nomeArquivo, setNomeArquivo] = useState('Base padrão (planilha enviada)');

  const handleFileChange = (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      setNomeArquivo(`Base carregada: ${file.name}`);
      alert(`Planilha "${file.name}" selecionada com sucesso!`);
    }
  };

  // --- FUNÇÃO EXCLUSIVA DE EXPORTAÇÃO DO DASHBOARD DE REQUIOSIÇÕES ---
  const exportarParaExcel = () => {
    const wb = XLSX.utils.book_new();

    // 1. Aba: KPIs Executivos
    const kpisData = [
      { Indicador: 'Volume de Requisições', Valor: '232 RCs', Detalhamento: '390 itens solicitados' },
      { Indicador: 'Lead Time Médio', Valor: '24,3 dias', Detalhamento: 'Pico de 85 dias registrados' },
      { Indicador: 'Pendências Críticas', Valor: '32,8%', Detalhamento: '76 RCs > 30 dias (18 > 60 dias)' },
      { Indicador: 'Cobertura Operacional', Valor: '40 centros', Detalhamento: '67.719 unidades solicitadas' }
    ];
    const wsKPIs = XLSX.utils.json_to_sheet(kpisData);
    XLSX.utils.book_append_sheet(wb, wsKPIs, 'KPIs Executivos');

    // 2. Aba: Evolução Temporal
    const wsEvolucao = XLSX.utils.json_to_sheet(evolucaoData);
    XLSX.utils.book_append_sheet(wb, wsEvolucao, 'Evolução Temporal');

    // 3. Aba: Faixas de Aging
    const wsAging = XLSX.utils.json_to_sheet(agingBarData);
    XLSX.utils.book_append_sheet(wb, wsAging, 'Faixas de Aging');

    // 4. Aba: Matriz de Priorização por Planta
    const wsMatriz = XLSX.utils.json_to_sheet(matrizPlantas);
    XLSX.utils.book_append_sheet(wb, wsMatriz, 'Matriz por Planta');

    // Gera o download do arquivo .xlsx
    XLSX.writeFile(wb, 'Dashboard_Requisicoes_Procurement.xlsx');
  };

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
          <p className="text-xs text-slate-400 mt-0.5 font-medium">
            {nomeArquivo}
          </p>
        </div>

        {/* BOTÕES DE AÇÃO */}
        <div className="flex items-center gap-3">
          <button 
            onClick={exportarParaExcel}
            className="px-4 py-2 bg-[#1E293B] hover:bg-[#334155] border border-slate-700 text-xs font-bold text-white rounded-lg transition inline-flex items-center gap-2 select-none shadow-sm cursor-pointer"
          >
            <svg className="w-4 h-4 text-[#00E599]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            Exportar dados (.xlsx)
          </button>

          <label className="px-4 py-2 bg-[#00E599] hover:bg-[#00C282] text-xs font-bold text-[#0A0E17] rounded-lg transition shadow-lg shadow-[#00E599]/10 cursor-pointer inline-flex items-center gap-2 select-none">
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
            </svg>
            Atualizar base (.xlsx)
            <input 
              type="file" 
              onChange={handleFileChange} 
              accept=".xlsx, .xls, .csv" 
              className="hidden" 
            />
          </label>
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
      </div>

      {/* AGING E CUMPRIMENTO DE DATA */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
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
        </div>

        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">CUMPRIMENTO DE DATA DE REMESSA</h2>
          <p className="text-[11px] text-slate-400 mb-2">Comparação entre Modificado em e Remessas (dia útil)</p>
          
          <div className="h-56">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={remessaData} innerRadius={65} outerRadius={85} paddingAngle={4} dataKey="value">
                  {remessaData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
                <Tooltip contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

    </div>
  );
}