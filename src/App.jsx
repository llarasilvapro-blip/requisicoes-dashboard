import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  BarChart, Bar, LabelList, PieChart, Pie, Cell, Legend
} from 'recharts';

// ------------------------------------------------------------------
// HELPERS DE EXTRAÇÃO / CLASSIFICAÇÃO
// ------------------------------------------------------------------
const extrairComprador = (acao) => {
  if (!acao) return 'Não atribuído';
  const acaoStr = String(acao);
  if (acaoStr.includes(':')) {
    return acaoStr.split(':')[1].trim();
  }
  return acaoStr;
};

const calcularFaixaAging = (dias) => {
  const d = Number(dias) || 0;
  if (d <= 15) return '0 - 15 dias';
  if (d <= 30) return '16 - 30 dias';
  if (d <= 60) return '31 - 60 dias';
  return 'Acima de 60 dias';
};

const normalizarGC = (gc) => {
  const s = String(gc || '').trim();
  if (!s || s === '-' || s === 'nan' || s === 'null') return 'Não informado';
  return s;
};

const normalizarObs = (obs) => {
  if (!obs || obs === 'nan' || obs === 'null') return null;
  return String(obs).trim().replace(/\s+/g, ' ').toUpperCase();
};

const parseData = (str) => {
  if (!str || str === 'nan' || str === 'null') return null;
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
};

const HOJE = new Date('2026-09-16');

export default function App() {
  const [baseDados, setBaseDados] = useState([]);
  const [filtroCentro, setFiltroCentro] = useState('Todos');
  const [filtroComprador, setFiltroComprador] = useState('Todos');
  const [filtroGC, setFiltroGC] = useState('Todos');
  const [filtroAging, setFiltroAging] = useState('Todos');
  const [nomeArquivo, setNomeArquivo] = useState('Faça o upload do arquivo Excel para carregar o dashboard.');

  // PARSER DO UPLOAD EXCEL
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

          const jsonDados = XLSX.utils.sheet_to_json(worksheet, { raw: false });

          if (jsonDados && jsonDados.length > 0) {
            const linhasValidas = jsonDados.filter(row => row['Requisição de compra'] && !isNaN(Number(row['Requisição de compra'])));

            const dadosFormatados = linhasValidas.map((row, idx) => {
              const dias = Number(row['Dias RC']) || 0;
              return {
                id: row['Concat'] || row['Requisição de compra'] || idx + 1,
                rc: String(row['Requisição de compra'] || ''),
                itemRc: row['Item ReqC'] || null,
                diasRC: dias,
                centro: String(row['Nome Centro'] || row['Centro'] || 'Outros'),
                comprador: extrairComprador(row['Ação RC']),
                grupoCompras: normalizarGC(row['GC Compras']),
                faixaAging: calcularFaixaAging(dias),
                qtdSolicitada: Number(row['Qtd.solicitada']) || 1,
                unidadeMedida: row['Unidade de medida'] || 'UN',
                material: row['Material'] || null,
                textoBreve: row['Texto breve'] || null,
                dataSolicitacao: row['Data da solicitação'] || null,
                modificadoEm: row['Modificado em'] || null,
                remessa: row['Remessas (de/até)'] || null,
                tratativa: row['Tratativa'] || null,
                dataTratativa: row['data de tratativa'] || null,
                obs: row['OBS'] || null,
              };
            });

            setBaseDados(dadosFormatados);
            setFiltroCentro('Todos');
            setFiltroComprador('Todos');
            setFiltroGC('Todos');
            setFiltroAging('Todos');
            setNomeArquivo(`Base carregada: ${file.name} (${dadosFormatados.length} itens)`);
          } else {
            alert('A planilha importada está vazia.');
          }
        } catch (err) {
          console.error('Erro ao ler o arquivo Excel:', err);
          alert('Erro ao ler o arquivo Excel. Verifique a estrutura do arquivo.');
        }
      };

      reader.readAsArrayBuffer(file);
    }
  };

  // OPÇÕES DOS FILTROS
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

  // FILTRAGEM
  const dadosFiltrados = useMemo(() => {
    return baseDados.filter(item => {
      const matchCentro = filtroCentro === 'Todos' || item.centro === filtroCentro;
      const matchComprador = filtroComprador === 'Todos' || item.comprador === filtroComprador;
      const matchGC = filtroGC === 'Todos' || item.grupoCompras === filtroGC;
      const matchAging = filtroAging === 'Todos' || item.faixaAging === filtroAging;
      return matchCentro && matchComprador && matchGC && matchAging;
    });
  }, [baseDados, filtroCentro, filtroComprador, filtroGC, filtroAging]);

  // KPIS
  const totalRCsUnicas = useMemo(() => new Set(dadosFiltrados.map(d => d.rc)).size, [dadosFiltrados]);

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

  // REMESSAS
  const remessaDataReal = useMemo(() => {
    let atrasada = 0, noPrazo = 0, semData = 0;
    dadosFiltrados.forEach(d => {
      const r = parseData(d.remessa);
      if (!r) { semData += 1; return; }
      if (r < HOJE) atrasada += 1; else noPrazo += 1;
    });
    return { atrasada, noPrazo, semData };
  }, [dadosFiltrados]);

  const percRemessaVencida = useMemo(() => {
    const total = dadosFiltrados.length;
    if (!total) return '0,0';
    return ((remessaDataReal.atrasada / total) * 100).toFixed(1).replace('.', ',');
  }, [dadosFiltrados, remessaDataReal]);

  // STATUS TRATATIVA
  const tratativaStatusData = useMemo(() => {
    const comTratativa = dadosFiltrados.filter(d => d.tratativa && String(d.tratativa).trim() !== '').length;
    const semTratativa = dadosFiltrados.length - comTratativa;
    return [
      { name: 'Sem tratativa registrada', value: semTratativa, color: '#f59e0b' },
      { name: 'Com tratativa definida', value: comTratativa, color: '#06b6d4' },
    ];
  }, [dadosFiltrados]);

  const percSemTratativa = useMemo(() => {
    const total = dadosFiltrados.length;
    if (!total) return '0,0';
    const sem = tratativaStatusData.find(t => t.name.startsWith('Sem'))?.value || 0;
    return ((sem / total) * 100).toFixed(1).replace('.', ',');
  }, [dadosFiltrados, tratativaStatusData]);

  // DADOS DE GRÁFICOS
  const concentracaoGCData = useMemo(() => {
    const contagem = {};
    dadosFiltrados.forEach(d => {
      contagem[d.grupoCompras] = (contagem[d.grupoCompras] || 0) + 1;
    });
    const total = dadosFiltrados.length || 1;
    return Object.entries(contagem)
      .map(([gc, qtd]) => ({ gc, qtd, perc: ((qtd / total) * 100).toFixed(1) }))
      .sort((a, b) => b.qtd - a.qtd);
  }, [dadosFiltrados]);

  const datasEvolucaoUnicas = useMemo(() => {
    const setDatas = new Set();
    dadosFiltrados.forEach(d => {
      const dt = parseData(d.dataSolicitacao);
      if (dt) setDatas.add(dt.toLocaleDateString('pt-BR'));
    });
    return Array.from(setDatas).sort().slice(-3);
  }, [dadosFiltrados]);

  const evolucaoDataLaurence = useMemo(() => {
    const mapa = {};
    dadosFiltrados.forEach(d => {
      const gc = d.grupoCompras || 'Não informado';
      const dt = parseData(d.dataSolicitacao);
      const dataStr = dt ? dt.toLocaleDateString('pt-BR') : 'Sem Data';

      if (!mapa[gc]) mapa[gc] = { gc, total: 0 };
      mapa[gc][dataStr] = (mapa[gc][dataStr] || 0) + 1;
      mapa[gc].total += 1;
    });

    return Object.values(mapa)
      .sort((a, b) => b.total - a.total)
      .slice(0, 8);
  }, [dadosFiltrados]);

  const agingBarDataDinamico = useMemo(() => {
    const faixas = ['0 - 15 dias', '16 - 30 dias', '31 - 60 dias', 'Acima de 60 dias'];
    const total = dadosFiltrados.length || 1;
    return faixas.map(faixa => {
      const itensNaFaixa = dadosFiltrados.filter(d => d.faixaAging === faixa);
      return {
        faixa,
        itens: itensNaFaixa.length,
        rcs: new Set(itensNaFaixa.map(d => d.rc)).size,
        percItens: ((itensNaFaixa.length / total) * 100).toFixed(1)
      };
    });
  }, [dadosFiltrados]);

  const analiseRemessaData = useMemo(() => {
    return [
      { status: 'No Prazo', qtd: remessaDataReal.noPrazo, fill: '#10b981' },
      { status: 'Vencida', qtd: remessaDataReal.atrasada, fill: '#ef4444' },
      { status: 'Sem Data', qtd: remessaDataReal.semData, fill: '#6b7280' },
    ];
  }, [remessaDataReal]);

  const motivosOBSData = useMemo(() => {
    const contagem = {};
    dadosFiltrados.forEach(d => {
      const obs = normalizarObs(d.obs);
      if (!obs) return;
      contagem[obs] = (contagem[obs] || 0) + 1;
    });
    const total = dadosFiltrados.length || 1;

    return Object.entries(contagem)
      .map(([motivo, qtd]) => ({ 
        motivo: motivo.length > 22 ? motivo.substring(0, 22) + '...' : motivo, 
        qtd,
        perc: ((qtd / total) * 100).toFixed(1)
      }))
      .sort((a, b) => b.qtd - a.qtd)
      .slice(0, 7);
  }, [dadosFiltrados]);

  const tratativaPieData = useMemo(() => {
    const contagem = {};
    dadosFiltrados.forEach(d => {
      const t = d.tratativa && String(d.tratativa).trim() !== '' ? String(d.tratativa).trim() : 'Sem Tratativa Registrada';
      contagem[t] = (contagem[t] || 0) + 1;
    });

    const total = dadosFiltrados.length || 1;
    return Object.entries(contagem).map(([name, value]) => ({
      name: name.length > 22 ? name.substring(0, 22) + '...' : name,
      value,
      perc: ((value / total) * 100).toFixed(1)
    })).sort((a, b) => b.value - a.value).slice(0, 7);
  }, [dadosFiltrados]);

  const rankingCompradores = useMemo(() => {
    const grupos = {};
    const totalItensGlobal = dadosFiltrados.length || 1;
    dadosFiltrados.forEach(d => {
      if (!grupos[d.comprador]) {
        grupos[d.comprador] = { comprador: d.comprador, itens: 0, rcsSet: new Set(), somaDias: 0, criticas60: 0 };
      }
      const g = grupos[d.comprador];
      g.itens += 1;
      g.rcsSet.add(d.rc);
      g.somaDias += d.diasRC;
      if (d.diasRC > 60) g.criticas60 += 1;
    });
    return Object.values(grupos)
      .map(g => ({
        comprador: g.comprador,
        rcs: g.rcsSet.size,
        itens: g.itens,
        pctItens: ((g.itens / totalItensGlobal) * 100).toFixed(1),
        agingMedio: (g.somaDias / g.itens).toFixed(1).replace('.', ','),
        criticas60: g.criticas60,
      }))
      .sort((a, b) => b.itens - a.itens);
  }, [dadosFiltrados]);

  const rankingCentros = useMemo(() => {
    const grupos = {};
    const total = dadosFiltrados.length || 1;
    dadosFiltrados.forEach(d => {
      const nomeLimpo = d.centro.replace('Indústria - ', '');
      if (!grupos[nomeLimpo]) {
        grupos[nomeLimpo] = { centro: nomeLimpo, itens: 0, rcsSet: new Set(), somaDias: 0 };
      }
      const g = grupos[nomeLimpo];
      g.itens += 1;
      g.rcsSet.add(d.rc);
      g.somaDias += d.diasRC;
    });
    return Object.values(grupos)
      .map(g => ({
        centro: g.centro.length > 20 ? g.centro.substring(0, 20) + '...' : g.centro,
        rcs: g.rcsSet.size,
        itens: g.itens,
        perc: ((g.itens / total) * 100).toFixed(1),
        agingMedio: Number((g.somaDias / g.itens).toFixed(1)),
      }))
      .sort((a, b) => b.itens - a.itens)
      .slice(0, 8);
  }, [dadosFiltrados]);
  return (
    <div className="min-h-screen bg-[#0b0f19] text-white p-6 font-sans">

      {/* CABEÇALHO */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 pb-4 border-b border-gray-800 gap-4">
        <div>
          <span className="text-[10px] text-cyan-400 font-semibold tracking-wider uppercase">PROCUREMENT · REQUISIÇÕES DE COMPRA</span>
          <h1 className="text-xl font-bold mt-0.5">Dashboard de Requisições</h1>
          <p className="text-xs text-gray-400 mt-0.5 font-medium">{nomeArquivo}</p>
        </div>

        <label className="flex items-center gap-2 bg-cyan-500 hover:bg-cyan-600 text-slate-950 text-xs font-semibold px-3 py-1.5 rounded cursor-pointer transition shadow-md select-none">
          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
          </svg>
          Atualizar base (.xlsx)
          <input type="file" onChange={handleFileChange} accept=".xlsx, .xls, .csv" className="hidden" />
        </label>
      </header>

      {/* FILTROS */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-3 mb-6 bg-[#131927] p-3 rounded-lg border border-gray-800 text-xs">
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Centro:</label>
          <select value={filtroCentro} onChange={(e) => setFiltroCentro(e.target.value)}
            className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white outline-none focus:border-cyan-400">
            {opcoesFiltros.centros.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Comprador:</label>
          <select value={filtroComprador} onChange={(e) => setFiltroComprador(e.target.value)}
            className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white outline-none focus:border-cyan-400">
            {opcoesFiltros.compradores.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Grupo de Compras:</label>
          <select value={filtroGC} onChange={(e) => setFiltroGC(e.target.value)}
            className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white outline-none focus:border-cyan-400">
            {opcoesFiltros.grupos.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Faixa de Aging:</label>
          <select value={filtroAging} onChange={(e) => setFiltroAging(e.target.value)}
            className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white outline-none focus:border-cyan-400">
            {opcoesFiltros.agings.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
          </select>
        </div>
      </div>

      {/* KPIS REORGANIZADOS E ESTILIZADOS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-4 mb-6">
        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">VOLUME REQUISIÇÕES</span>
          <p className="text-2xl font-bold text-cyan-400 mt-1">{totalRCsUnicas} RCs</p>
          <span className="text-[10px] text-gray-500">{dadosFiltrados.length} itens solicitados</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">LEAD TIME MÉDIO</span>
          <p className="text-2xl font-bold text-white mt-1">{leadTimeMedio} dias</p>
          <span className="text-[10px] text-amber-400 font-medium">Pico: {picoDias} dias</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">PENDÊNCIAS CRÍTICAS</span>
          <p className="text-2xl font-bold text-red-400 mt-1">{pendenciasCriticas.perc}%</p>
          <span className="text-[10px] text-gray-500">{pendenciasCriticas.rcs30} &gt; 30d · {pendenciasCriticas.rcs60} &gt; 60d</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">REMESSA VENCIDA</span>
          <p className="text-2xl font-bold text-red-400 mt-1">{percRemessaVencida}%</p>
          <span className="text-[10px] text-gray-500">{remessaDataReal.atrasada} itens atrasados</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">SEM TRATATIVA</span>
          <p className="text-2xl font-bold text-amber-400 mt-1">{percSemTratativa}%</p>
          <span className="text-[10px] text-gray-500">Sem ação registrada</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">COBERTURA OPERACIONAL</span>
          <p className="text-2xl font-bold text-emerald-400 mt-1">{coberturaOperacional.centrosAtivos} centros</p>
          <span className="text-[10px] text-gray-500">{coberturaOperacional.qtdTotalUnidades.toLocaleString('pt-BR')} un. solicitadas</span>
        </div>
      </div>

      {baseDados.length > 0 ? (
        <>
          {/* LINHA 1: RANKING COMPRADORES E FAIXA DE AGING */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6 text-xs">
            {/* 1. RANKING COMPRADORES */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">1. RANKING DE COMPRADORES (CARGA x AGING)</h2>
              <div className="overflow-x-auto max-h-64">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 bg-[#131927]">
                    <tr className="border-b border-gray-800 text-gray-400 text-[10px] uppercase">
                      <th className="pb-2">Comprador</th>
                      <th className="pb-2 text-center">RCs</th>
                      <th className="pb-2 text-center">Itens (% Tot)</th>
                      <th className="pb-2 text-right">Aging Médio</th>
                      <th className="pb-2 text-right">&gt; 60 dias</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800/50">
                    {rankingCompradores.map((r, i) => (
                      <tr key={i} className="hover:bg-[#1a2234]">
                        <td className="py-2 text-gray-200 font-medium">{r.comprador}</td>
                        <td className="py-2 text-center text-cyan-400 font-semibold">{r.rcs}</td>
                        <td className="py-2 text-center text-gray-300">
                          {r.itens} <span className="text-[10px] text-cyan-400">({r.pctItens}%)</span>
                        </td>
                        <td className="py-2 text-right text-gray-300">{r.agingMedio} d</td>
                        <td className={`py-2 text-right font-bold ${r.criticas60 > 0 ? 'text-red-400' : 'text-emerald-400'}`}>
                          {r.criticas60}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* 2. FAIXA DE AGING */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">2. DISTRIBUIÇÃO POR FAIXA DE AGING</h2>
              <div className="h-60">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={agingBarDataDinamico}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" vertical={false} />
                    <XAxis dataKey="faixa" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <YAxis stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <Tooltip 
                      formatter={(val, name, props) => name === "Itens" 
                        ? [`${val} (${props.payload.percItens}%)`, name] 
                        : [val, name]}
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} 
                    />
                    <Bar dataKey="itens" fill="#06b6d4" radius={[2, 2, 0, 0]} name="Itens">
                      <LabelList dataKey="percItens" position="top" formatter={(v) => `${v}%`} fill="#9ca3af" fontSize={10} />
                    </Bar>
                    <Bar dataKey="rcs" fill="#f59e0b" radius={[2, 2, 0, 0]} name="RCs" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* LINHA 2: TRATATIVAS E DATAS DE REMESSA (COM GRÁFICO DONUT REVISADO) */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
            {/* 3. TRATATIVA */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">3. DISTRIBUIÇÃO DOS STATUS DE TRATATIVA</h2>
              <div className="h-60">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={tratativaPieData} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" horizontal={false} />
                    <XAxis type="number" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <YAxis type="category" dataKey="name" stroke="#6b7280" width={170} tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <Tooltip 
                      formatter={(value, name, props) => [`${value} itens (${props.payload.perc}%)`, 'Volume']}
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} 
                    />
                    <Bar dataKey="value" fill="#10b981" radius={[0, 4, 4, 0]} name="Itens">
                      <LabelList dataKey="perc" position="right" formatter={(v) => `${v}%`} fill="#9ca3af" fontSize={10} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* 4. ANÁLISE REMESSA (DONUT MELHORADO) */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">4. ANÁLISE DAS DATAS DE REMESSA</h2>
              <div className="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center h-60">
                <div className="sm:col-span-6 h-full relative">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={analiseRemessaData.filter(d => d.qtd > 0)}
                        dataKey="qtd"
                        nameKey="status"
                        cx="50%"
                        cy="50%"
                        outerRadius={70}
                        innerRadius={45}
                        paddingAngle={3}
                      >
                        {analiseRemessaData.filter(d => d.qtd > 0).map((entry, index) => (
                          <Cell key={`cell-remessa-${index}`} fill={entry.fill} />
                        ))}
                      </Pie>
                      <Tooltip
                        contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }}
                        formatter={(value, name) => [
                          `${value} itens (${((value / dadosFiltrados.length) * 100).toFixed(1)}%)`,
                          name
                        ]}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </div>

                <div className="sm:col-span-6 space-y-2">
                  {analiseRemessaData.map((item, idx) => {
                    const pct = dadosFiltrados.length ? ((item.qtd / dadosFiltrados.length) * 100).toFixed(1) : 0;
                    return (
                      <div key={idx} className="flex justify-between items-center p-2 rounded bg-[#1a2234] border border-gray-800/80 text-xs">
                        <div className="flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: item.fill }}></span>
                          <span className="text-gray-300 font-medium">{item.status}</span>
                        </div>
                        <div className="text-right">
                          <span className="font-bold text-white">{item.qtd}</span>
                          <span className="text-[10px] text-gray-400 ml-1">({pct}%)</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>

          {/* LINHA 3: MOTIVOS E GRUPOS DE COMPRAS */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
            {/* 5. OBS MOTIVOS */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">5. PRINCIPAIS MOTIVOS DE PENDÊNCIA (OBS)</h2>
              <div className="h-60">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={motivosOBSData} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" horizontal={false} />
                    <XAxis type="number" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <YAxis type="category" dataKey="motivo" stroke="#6b7280" width={170} tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <Tooltip 
                      formatter={(val, name, props) => [`${val} ocorrências (${props.payload.perc}%)`, 'Ocorrências']}
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} 
                    />
                    <Bar dataKey="qtd" fill="#f59e0b" radius={[0, 4, 4, 0]} name="Ocorrências">
                      <LabelList dataKey="perc" position="right" formatter={(v) => `${v}%`} fill="#9ca3af" fontSize={10} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* 6. CONCENTRAÇÃO GC */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">6. CONCENTRAÇÃO POR GRUPO DE COMPRAS (GC)</h2>
              <div className="h-60">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={concentracaoGCData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" vertical={false} />
                    <XAxis dataKey="gc" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <YAxis stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <Tooltip
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }}
                      formatter={(value, name, props) => [`${value} itens (${props.payload.perc}%)`, 'Volume']}
                    />
                    <Bar dataKey="qtd" fill="#06b6d4" radius={[2, 2, 0, 0]} name="Itens">
                      <LabelList dataKey="perc" position="top" formatter={(v) => `${v}%`} fill="#9ca3af" fontSize={10} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* LINHA 4: EVOLUÇÃO TEMPORAL E TOP CENTROS */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
            {/* 7. EVOLUÇÃO TEMPORAL */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">7. EVOLUÇÃO TEMPORAL POR GRUPO DE COMPRAS</h2>
              <div className="h-60">
                {evolucaoDataLaurence.length > 0 ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={evolucaoDataLaurence}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" vertical={false} />
                      <XAxis dataKey="gc" stroke="#6b7280" tick={{ fontSize: 9, fill: '#9ca3af' }} angle={-20} textAnchor="end" />
                      <YAxis stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                      <Tooltip contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} />
                      {datasEvolucaoUnicas.map((dtStr, idx) => (
                        <Bar
                          key={dtStr}
                          dataKey={dtStr}
                          name={dtStr}
                          fill={idx === 0 ? '#06b6d4' : idx === 1 ? '#6366f1' : '#f59e0b'}
                          radius={[2, 2, 0, 0]}
                        />
                      ))}
                    </BarChart>
                  </ResponsiveContainer>
                ) : (
                  <div className="h-full flex items-center justify-center text-xs text-gray-500">
                    Carregue a planilha para visualizar a evolução temporal.
                  </div>
                )}
              </div>
            </div>

            {/* 8. TOP CENTROS (EIXO Y EXPANDIDO) */}
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">8. TOP CENTROS COM MAIS ITENS PENDENTES</h2>
              <div className="h-60">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={rankingCentros} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" horizontal={false} />
                    <XAxis type="number" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                    <YAxis type="category" dataKey="centro" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} width={170} />
                    <Tooltip
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }}
                      formatter={(value, name, props) => name === 'itens' ? [`${value} itens (${props.payload.perc}%) · aging médio ${props.payload.agingMedio}d`, 'Volume'] : [value, name]}
                    />
                    <Bar dataKey="itens" fill="#10b981" radius={[0, 4, 4, 0]} name="Itens">
                      <LabelList dataKey="perc" position="right" formatter={(v) => `${v}%`} fill="#9ca3af" fontSize={10} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        </>
      ) : (
        <div className="flex flex-col items-center justify-center py-24 bg-[#131927] rounded-lg border border-gray-800 text-gray-500 gap-3">
          <svg className="w-10 h-10 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
          </svg>
          <p className="text-xs font-medium">Carregue o arquivo de requisições (.xlsx) no botão superior para visualizar o dashboard.</p>
        </div>
      )}
    </div>
  );
}