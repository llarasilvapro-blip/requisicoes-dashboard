import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  BarChart, Bar, PieChart, Pie, Cell 
} from 'recharts';

// --- ESTRUTURA INICIAL BASEADA NA SUA PLANILHA ---
const evolucaoDataMock = [
  { mes: '2020-05', solicitacoes: 18, modificacoes: 5 },
  { mes: '2020-06', solicitacoes: 65, modificacoes: 12 },
  { mes: '2020-07', solicitacoes: 125, modificacoes: 30 },
  { mes: '2020-08', solicitacoes: 235, modificacoes: 310 },
  { mes: '2020-09', solicitacoes: 40, modificacoes: 45 },
];

const remessaData = [
  { name: 'Remessa no prazo', value: 82, color: '#00E599' },
  { name: 'Remessa atrasada', value: 12, color: '#F59E0B' },
  { name: 'Sem data', value: 6, color: '#A855F7' },
];

export default function App() {
  const [baseDados, setBaseDados] = useState([]);
  const [filtroCentro, setFiltroCentro] = useState('Todos');
  const [filtroComprador, setFiltroComprador] = useState('Todos');
  const [filtroGC, setFiltroGC] = useState('Todos');
  const [filtroAging, setFiltroAging] = useState('Todos');
  const [nomeArquivo, setNomeArquivo] = useState('Base padrão (Requisições procurement.xlsx)');

  // HELPER PARA CLASSIFICAR FAIXA DE AGING BASEADO EM 'Dias RC'
  const calcularFaixaAging = (dias) => {
    const d = Number(dias) || 0;
    if (d <= 15) return '0 - 15 dias';
    if (d <= 30) return '16 - 30 dias';
    if (d <= 60) return '31 - 60 dias';
    return 'Acima de 60 dias';
  };

  // HELPER PARA EXTRAIR O NOME DO COMPRADOR DA COLUNA 'Ação RC'
  const extrairComprador = (acao) => {
    if (!acao) return 'Não atribuído';
    const acaoStr = String(acao);
    if (acaoStr.includes(':')) {
      return acaoStr.split(':')[1].trim();
    }
    return acaoStr;
  };

  // LEITURA E MAPEAMENTO EXATO DA SUA PLANILHA EXCEL
  const handleFileChange = (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          
          const jsonDados = XLSX.utils.sheet_to_json(worksheet);

          if (jsonDados && jsonDados.length > 0) {
            const dadosFormatados = jsonDados.map((row, idx) => {
              const dias = Number(row['Dias RC']) || 0;
              return {
                id: row['Concat'] || row['Requisição de compra'] || idx + 1,
                rc: row['Requisição de compra'],
                diasRC: dias,
                centro: String(row['Nome Centro'] || row['Centro'] || 'Outros'),
                comprador: extrairComprador(row['Ação RC']),
                grupoCompras: String(row['GC Compras'] || 'Outros'),
                faixaAging: calcularFaixaAging(dias),
                qtdSolicitada: Number(row['Qtd.solicitada']) || 1
              };
            });

            setBaseDados(dadosFormatados);
            
            // Reseta filtros ao carregar nova base
            setFiltroCentro('Todos');
            setFiltroComprador('Todos');
            setFiltroGC('Todos');
            setFiltroAging('Todos');

            setNomeArquivo(`Base carregada: ${file.name}`);
          } else {
            alert('A planilha importada está vazia.');
          }
        } catch (err) {
          console.error("Erro ao ler o arquivo Excel:", err);
          alert("Erro ao ler o arquivo Excel. Verifique a estrutura do arquivo.");
        }
      };

      reader.readAsArrayBuffer(file);
    }
  };

  // DROPDOWNS GERADOS DINAMICAMENTE
  const opcoesFiltros = useMemo(() => {
    const getUnicos = (key) => {
      const vals = baseDados.map(item => item[key]).filter(Boolean);
      return ['Todos', ...Array.from(new Set(vals)).sort()];
    };

    return {
      centros: getUnicos('centro'),
      compradores: getUnicos('comprador'),
      grupos: getUnicos('grupoCompras'),
      agings: ['Todos', '0 - 15 dias', '16 - 30 dias', '31 - 60 dias', 'Acima de 60 dias'],
    };
  }, [baseDados]);

  // FILTRAGEM DOS DADOS
  const dadosFiltrados = useMemo(() => {
    return baseDados.filter(item => {
      const matchCentro = filtroCentro === 'Todos' || item.centro === filtroCentro;
      const matchComprador = filtroComprador === 'Todos' || item.comprador === filtroComprador;
      const matchGC = filtroGC === 'Todos' || item.grupoCompras === filtroGC;
      const matchAging = filtroAging === 'Todos' || item.faixaAging === filtroAging;
      return matchCentro && matchComprador && matchGC && matchAging;
    });
  }, [baseDados, filtroCentro, filtroComprador, filtroGC, filtroAging]);

  // KPIS DINÂMICOS
  const totalRCsUnicas = useMemo(() => {
    return new Set(dadosFiltrados.map(d => d.rc)).size;
  }, [dadosFiltrados]);

  const leadTimeMedio = useMemo(() => {
    if (!dadosFiltrados.length) return '0,0';
    const soma = dadosFiltrados.reduce((acc, curr) => acc + curr.diasRC, 0);
    return (soma / dadosFiltrados.length).toFixed(1).replace('.', ',');
  }, [dadosFiltrados]);

  const picoDias = useMemo(() => {
    if (!dadosFiltrados.length) return 0;
    return Math.max(...dadosFiltrados.map(d => d.diasRC));
  }, [dadosFiltrados]);

  const pendenciasCriticas = useMemo(() => {
    if (!dadosFiltrados.length) return { perc: '0,0', rcs30: 0, rcs60: 0 };
    const rcs30 = dadosFiltrados.filter(d => d.diasRC > 30).length;
    const rcs60 = dadosFiltrados.filter(d => d.diasRC > 60).length;
    const perc = ((rcs30 / dadosFiltrados.length) * 100).toFixed(1).replace('.', ',');
    return { perc, rcs30, rcs60 };
  }, [dadosFiltrados]);

  const coberturaOperacional = useMemo(() => {
    const centrosAtivos = new Set(dadosFiltrados.map(d => d.centro)).size;
    const qtdTotalUnidades = dadosFiltrados.reduce((acc, curr) => acc + curr.qtdSolicitada, 0);
    return { centrosAtivos, qtdTotalUnidades };
  }, [dadosFiltrados]);

  // GRÁFICO BARRAS DE AGING REATIVO
  const agingBarDataDinamico = useMemo(() => {
    const faixas = ['0 - 15 dias', '16 - 30 dias', '31 - 60 dias', 'Acima de 60 dias'];
    return faixas.map(faixa => {
      const itensNaFaixa = dadosFiltrados.filter(d => d.faixaAging === faixa);
      const qtdItens = itensNaFaixa.length;
      const qtdRCs = new Set(itensNaFaixa.map(d => d.rc)).size;
      return { faixa, itens: qtdItens, rcs: qtdRCs };
    });
  }, [dadosFiltrados]);

  return (
    <div className="min-h-screen bg-[#0A0E17] text-[#E2E8F0] p-6 font-sans">
      
      {/* CABEÇALHO */}
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

        {/* BOTÃO DE UPLOAD (BOTÃO DE EXPORTAR REMOVIDO) */}
        <div className="flex items-center gap-3">
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
          
          {/* Centro */}
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Centro</label>
            <select 
              value={filtroCentro} 
              onChange={(e) => setFiltroCentro(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              {opcoesFiltros.centros.map(opcao => (
                <option key={opcao} value={opcao}>{opcao}</option>
              ))}
            </select>
          </div>

          {/* Comprador */}
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Comprador</label>
            <select 
              value={filtroComprador} 
              onChange={(e) => setFiltroComprador(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              {opcoesFiltros.compradores.map(opcao => (
                <option key={opcao} value={opcao}>{opcao}</option>
              ))}
            </select>
          </div>

          {/* Grupo de Compras */}
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Grupo de compras</label>
            <select 
              value={filtroGC} 
              onChange={(e) => setFiltroGC(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              {opcoesFiltros.grupos.map(opcao => (
                <option key={opcao} value={opcao}>{opcao}</option>
              ))}
            </select>
          </div>

          {/* Faixa de Aging */}
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Faixa de aging</label>
            <select 
              value={filtroAging} 
              onChange={(e) => setFiltroAging(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]"
            >
              {opcoesFiltros.agings.map(opcao => (
                <option key={opcao} value={opcao}>{opcao}</option>
              ))}
            </select>
          </div>
        </div>

        {/* CONTAGEM DINÂMICA DE ITENS */}
        <div className="text-right">
          <span className="text-xs font-semibold text-slate-400">
            {dadosFiltrados.length} de {baseDados.length} itens
          </span>
        </div>
      </div>

      {/* CARDS DE KPIS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">VOLUME DE REQUIÇÕES</p>
          <h3 className="text-2xl font-black text-[#00E599] mt-1">{totalRCsUnicas} RCs</h3>
          <p className="text-xs text-slate-400 mt-1">{dadosFiltrados.length} itens solicitados</p>
        </div>

        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">LEAD TIME MÉDIO (DIAS RC)</p>
          <h3 className="text-2xl font-black text-white mt-1">{leadTimeMedio} dias</h3>
          <p className="text-xs text-slate-400 mt-1">Pico de {picoDias} dias</p>
        </div>

        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">PENDÊNCIAS CRÍTICAS</p>
          <h3 className="text-2xl font-black text-white mt-1">{pendenciasCriticas.perc}%</h3>
          <p className="text-xs text-slate-400 mt-1">
            {pendenciasCriticas.rcs30} RCs &gt; 30 dias · {pendenciasCriticas.rcs60} &gt; 60 dias
          </p>
        </div>

        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">COBERTURA OPERACIONAL</p>
          <h3 className="text-2xl font-black text-white mt-1">{coberturaOperacional.centrosAtivos} centros</h3>
          <p className="text-xs text-slate-400 mt-1">{coberturaOperacional.qtdTotalUnidades.toLocaleString('pt-BR')} unidades solicitadas</p>
        </div>
      </div>

      {/* EVOLUÇÃO TEMPORAL */}
      <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80 mb-6">
        <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">EVOLUÇÃO TEMPORAL</h2>
        <p className="text-[11px] text-slate-400 mb-4">Solicitações abertas x modificações registradas por mês</p>
        
        <div className="h-64">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={evolucaoDataMock}>
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
              <BarChart data={agingBarDataDinamico}>
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