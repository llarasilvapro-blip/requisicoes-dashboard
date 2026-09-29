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

// PALETA DE CORES
const CORES_PIZZA = ['#00E599', '#3B82F6', '#F59E0B', '#EC4899', '#8B5CF6', '#14B8A6', '#6366F1'];

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

  // EVOLUÇÃO TEMPORAL (ESTILO LAURENCE)
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
      .slice(0, 12);
  }, [dadosFiltrados]);

  // FAIXAS DE AGING
  const agingBarDataDinamico = useMemo(() => {
    const faixas = ['0 - 15 dias', '16 - 30 dias', '31 - 60 dias', 'Acima de 60 dias'];
    return faixas.map(faixa => {
      const itensNaFaixa = dadosFiltrados.filter(d => d.faixaAging === faixa);
      return {
        faixa,
        itens: itensNaFaixa.length,
        rcs: new Set(itensNaFaixa.map(d => d.rc)).size,
      };
    });
  }, [dadosFiltrados]);

  // REMESSAS + MOTIVOS
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

  const motivosPendenciaEDataData = useMemo(() => {
    const contagem = {
      'REMESSA VENCIDA': remessaDataReal.atrasada,
      'REMESSA NO PRAZO': remessaDataReal.noPrazo,
      'SEM DATA REMESSA': remessaDataReal.semData,
    };
    
    dadosFiltrados.forEach(d => {
      const obs = normalizarObs(d.obs);
      if (!obs) return;
      contagem[obs] = (contagem[obs] || 0) + 1;
    });

    return Object.entries(contagem)
      .map(([motivo, qtd]) => ({ motivo, qtd }))
      .sort((a, b) => b.qtd - a.qtd)
      .slice(0, 10);
  }, [dadosFiltrados, remessaDataReal]);

  // STATUS TRATATIVA
  const tratativaStatusData = useMemo(() => {
    const comTratativa = dadosFiltrados.filter(d => d.tratativa && String(d.tratativa).trim() !== '').length;
    const semTratativa = dadosFiltrados.length - comTratativa;
    return [
      { name: 'Sem tratativa registrada', value: semTratativa, color: '#F59E0B' },
      { name: 'Com tratativa definida', value: comTratativa, color: '#00E599' },
    ];
  }, [dadosFiltrados]);

  const percSemTratativa = useMemo(() => {
    const total = dadosFiltrados.length;
    if (!total) return '0,0';
    const sem = tratativaStatusData.find(t => t.name.startsWith('Sem'))?.value || 0;
    return ((sem / total) * 100).toFixed(1).replace('.', ',');
  }, [dadosFiltrados, tratativaStatusData]);

  // 1. DADOS PARA O GRÁFICO DE PIZZA (STATUS DE TRATATIVA)
  const tratativaPieData = useMemo(() => {
    const contagem = {};
    dadosFiltrados.forEach(d => {
      const t = d.tratativa && String(d.tratativa).trim() !== '' ? String(d.tratativa).trim() : 'Sem Tratativa Registrada';
      contagem[t] = (contagem[t] || 0) + 1;
    });

    const total = dadosFiltrados.length || 1;
    return Object.entries(contagem).map(([name, value]) => ({
      name,
      value,
      perc: ((value / total) * 100).toFixed(1).replace('.', ',')
    })).sort((a, b) => b.value - a.value);
  }, [dadosFiltrados]);

  // 2. DADOS PARA O GRÁFICO DE ANÁLISE DE DATAS DE REMESSA
  const analiseRemessaData = useMemo(() => {
    return [
      { status: 'Remessa No Prazo', qtd: remessaDataReal.noPrazo, fill: '#00E599' },
      { status: 'Remessa Vencida', qtd: remessaDataReal.atrasada, fill: '#EF4444' },
      { status: 'Sem Data Registrada', qtd: remessaDataReal.semData, fill: '#64748B' },
    ];
  }, [remessaDataReal]);

  // CONCENTRAÇÃO GC
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

  // RANKINGS
  const rankingCompradores = useMemo(() => {
    const grupos = {};
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
        agingMedio: (g.somaDias / g.itens).toFixed(1).replace('.', ','),
        criticas60: g.criticas60,
      }))
      .sort((a, b) => b.itens - a.itens);
  }, [dadosFiltrados]);

  const rankingCentros = useMemo(() => {
    const grupos = {};
    dadosFiltrados.forEach(d => {
      if (!grupos[d.centro]) {
        grupos[d.centro] = { centro: d.centro, itens: 0, rcsSet: new Set(), somaDias: 0 };
      }
      const g = grupos[d.centro];
      g.itens += 1;
      g.rcsSet.add(d.rc);
      g.somaDias += d.diasRC;
    });
    return Object.values(grupos)
      .map(g => ({
        centro: g.centro,
        rcs: g.rcsSet.size,
        itens: g.itens,
        agingMedio: Number((g.somaDias / g.itens).toFixed(1)),
      }))
      .sort((a, b) => b.itens - a.itens)
      .slice(0, 8);
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

        <div className="flex items-center gap-3">
          <label className="px-4 py-2 bg-[#00E599] hover:bg-[#00C282] text-xs font-bold text-[#0A0E17] rounded-lg transition shadow-lg shadow-[#00E599]/10 cursor-pointer inline-flex items-center gap-2 select-none">
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
            </svg>
            Atualizar base (.xlsx)
            <input type="file" onChange={handleFileChange} accept=".xlsx, .xls, .csv" className="hidden" />
          </label>
        </div>
      </div>

      {/* BARRA DE FILTROS */}
      <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80 mb-6 flex flex-col md:flex-row justify-between items-center gap-4">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 w-full md:w-auto flex-1 max-w-4xl">
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Centro</label>
            <select value={filtroCentro} onChange={(e) => setFiltroCentro(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]">
              {opcoesFiltros.centros.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Comprador</label>
            <select value={filtroComprador} onChange={(e) => setFiltroComprador(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]">
              {opcoesFiltros.compradores.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Grupo de compras</label>
            <select value={filtroGC} onChange={(e) => setFiltroGC(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]">
              {opcoesFiltros.grupos.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-[10px] text-slate-400 mb-1 font-medium">Faixa de aging</label>
            <select value={filtroAging} onChange={(e) => setFiltroAging(e.target.value)}
              className="w-full bg-[#1A2333] border border-slate-700/60 rounded-lg text-xs text-white p-2 outline-none focus:border-[#00E599]">
              {opcoesFiltros.agings.map(opcao => <option key={opcao} value={opcao}>{opcao}</option>)}
            </select>
          </div>
        </div>
        <div className="text-right">
          <span className="text-xs font-semibold text-slate-400">
            {dadosFiltrados.length} de {baseDados.length} itens
          </span>
        </div>
      </div>

      {/* CARDS DE KPIS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">VOLUME DE REQUISIÇÕES</p>
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
          <p className="text-xs text-slate-400 mt-1">{pendenciasCriticas.rcs30} itens &gt; 30 dias · {pendenciasCriticas.rcs60} &gt; 60 dias</p>
        </div>
        <div className="bg-[#111726] p-4 rounded-xl border border-slate-800/80">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">COBERTURA OPERACIONAL</p>
          <h3 className="text-2xl font-black text-white mt-1">{coberturaOperacional.centrosAtivos} centros</h3>
          <p className="text-xs text-slate-400 mt-1">{coberturaOperacional.qtdTotalUnidades.toLocaleString('pt-BR')} unidades solicitadas</p>
        </div>
      </div>

      {/* KPIS DE RISCO */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
        <div className="bg-[#111726] p-4 rounded-xl border border-red-500/30">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">REMESSA JÁ VENCIDA (vs. hoje)</p>
          <h3 className="text-2xl font-black text-red-400 mt-1">{percRemessaVencida}%</h3>
          <p className="text-xs text-slate-400 mt-1">
            {remessaDataReal.atrasada} itens com data de remessa anterior a {HOJE.toLocaleDateString('pt-BR')}
          </p>
        </div>
        <div className="bg-[#111726] p-4 rounded-xl border border-amber-500/30">
          <p className="text-[10px] font-bold tracking-wider text-slate-400 uppercase">SEM TRATATIVA REGISTRADA</p>
          <h3 className="text-2xl font-black text-amber-400 mt-1">{percSemTratativa}%</h3>
          <p className="text-xs text-slate-400 mt-1">
            {tratativaStatusData.find(t => t.name.startsWith('Sem'))?.value || 0} itens aguardando ação/comentário do comprador
          </p>
        </div>
      </div>

      {/* EVOLUÇÃO TEMPORAL (LAURENCE) */}
      <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80 mb-6">
        <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">Qtde_Requisições por Gestão_Compras e Date</h2>
        <p className="text-[11px] text-slate-400 mb-4">Volume de requisições por grupo de compras comparado por data</p>
        <div className="h-72">
          {evolucaoDataLaurence.length > 0 ? (
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={evolucaoDataLaurence} margin={{ top: 20, right: 30, left: 0, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" vertical={false} />
                <XAxis dataKey="gc" stroke="#64748B" fontSize={10} angle={-25} textAnchor="end" />
                <YAxis stroke="#64748B" fontSize={10} />
                <Tooltip contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }} />
                {datasEvolucaoUnicas.map((dtStr, idx) => (
                  <Bar
                    key={dtStr}
                    dataKey={dtStr}
                    name={dtStr}
                    fill={idx === 0 ? '#3B82F6' : idx === 1 ? '#1D4ED8' : '#EA580C'}
                    radius={[3, 3, 0, 0]}
                  >
                    <LabelList dataKey={dtStr} position="top" fill="#94A3B8" fontSize={10} />
                  </Bar>
                ))}
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div className="h-full flex items-center justify-center text-xs text-slate-500">
              Carregue a planilha para visualizar os dados.
            </div>
          )}
        </div>
      </div>

      {/* AGING E PRINCIPAIS MOTIVOS */}
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
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">PRINCIPAIS MOTIVOS DE PENDÊNCIA (OBS + PRAZO REMESSA)</h2>
          <p className="text-[11px] text-slate-400 mb-4">Ocorrências por motivo de pendência e status do prazo de remessa</p>
          <div className="h-60">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={motivosPendenciaEDataData} layout="vertical" margin={{ left: 10, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" horizontal={false} />
                <XAxis type="number" stroke="#64748B" fontSize={10} allowDecimals={false} />
                <YAxis type="category" dataKey="motivo" stroke="#64748B" fontSize={9} width={150} />
                <Tooltip contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }} />
                <Bar dataKey="qtd" fill="#00E599" radius={[0, 3, 3, 0]} name="Ocorrências" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* CONCENTRAÇÃO POR GRUPO DE COMPRAS */}
      <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80 mb-6">
        <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">CONCENTRAÇÃO POR GRUPO DE COMPRAS (GC)</h2>
        <p className="text-[11px] text-slate-400 mb-4">Participação de cada grupo de compras no volume total de itens</p>
        <div className="h-56">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={concentracaoGCData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" vertical={false} />
              <XAxis dataKey="gc" stroke="#64748B" fontSize={10} />
              <YAxis stroke="#64748B" fontSize={10} allowDecimals={false} />
              <Tooltip
                contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }}
                formatter={(value, name, props) => [`${value} itens (${props.payload.perc}%)`, 'Volume']}
              />
              <Bar dataKey="qtd" fill="#00E599" radius={[3, 3, 0, 0]} name="Itens">
                <LabelList dataKey="perc" position="top" formatter={(v) => `${v}%`} fill="#94A3B8" fontSize={10} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* RANKING DE COMPRADORES E CENTROS */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">RANKING DE COMPRADORES</h2>
          <p className="text-[11px] text-slate-400 mb-3">Carga de trabalho e aging médio por comprador</p>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="text-slate-400 text-[10px] uppercase border-b border-slate-800">
                  <th className="text-left py-2 font-semibold">Comprador</th>
                  <th className="text-right py-2 font-semibold">RCs</th>
                  <th className="text-right py-2 font-semibold">Itens</th>
                  <th className="text-right py-2 font-semibold">Aging médio</th>
                  <th className="text-right py-2 font-semibold">&gt; 60 dias</th>
                </tr>
              </thead>
              <tbody>
                {rankingCompradores.length > 0 ? (
                  rankingCompradores.map(r => (
                    <tr key={r.comprador} className="border-b border-slate-800/60 last:border-0">
                      <td className="py-2 text-slate-200">{r.comprador}</td>
                      <td className="py-2 text-right text-slate-300">{r.rcs}</td>
                      <td className="py-2 text-right text-slate-300">{r.itens}</td>
                      <td className="py-2 text-right text-slate-300">{r.agingMedio} d</td>
                      <td className={`py-2 text-right font-semibold ${r.criticas60 > 0 ? 'text-red-400' : 'text-slate-300'}`}>{r.criticas60}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan="5" className="py-4 text-center text-slate-500">Nenhum dado carregado</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">TOP CENTROS COM MAIS ITENS PENDENTES</h2>
          <p className="text-[11px] text-slate-400 mb-4">8 centros com maior volume de requisições em aberto</p>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={rankingCentros} layout="vertical" margin={{ left: 10, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" horizontal={false} />
                <XAxis type="number" stroke="#64748B" fontSize={10} allowDecimals={false} />
                <YAxis type="category" dataKey="centro" stroke="#64748B" fontSize={9} width={150} />
                <Tooltip
                  contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }}
                  formatter={(value, name, props) => name === 'itens' ? [`${value} itens · aging médio ${props.payload.agingMedio}d`, 'Volume'] : [value, name]}
                />
                <Bar dataKey="itens" fill="#00E599" radius={[0, 3, 3, 0]} name="itens" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* SEÇÃO DUPLA FINAL: PIZZA DE TRATATIVAS + ANÁLISE DE DATAS DE REMESSA LADO A LADO */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
        
        {/* GRÁFICO DE PIZZA: STATUS DE TRATATIVA */}
        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">DISTRIBUIÇÃO DOS STATUS DE TRATATIVA</h2>
          <p className="text-[11px] text-slate-400 mb-2">Proporção total de itens por status de acompanhamento</p>
          <div className="h-64">
            {tratativaPieData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={tratativaPieData}
                    dataKey="value"
                    nameKey="name"
                    cx="50%"
                    cy="50%"
                    outerRadius={80}
                    innerRadius={45}
                    paddingAngle={2}
                  >
                    {tratativaPieData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={CORES_PIZZA[index % CORES_PIZZA.length]} />
                    ))}
                  </Pie>
                  <Tooltip
                    contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }}
                    formatter={(value, name, props) => [`${value} itens (${props.payload.perc}%)`, name]}
                  />
                  <Legend
                    layout="vertical"
                    verticalAlign="middle"
                    align="right"
                    wrapperStyle={{ fontSize: '10px', color: '#94A3B8' }}
                  />
                </PieChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-xs text-slate-500">
                Faça o upload do ficheiro Excel para carregar os dados.
              </div>
            )}
          </div>
        </div>

        {/* GRÁFICO DE BARRAS: ANÁLISE DAS DATAS DE REMESSA */}
        <div className="bg-[#111726] p-5 rounded-xl border border-slate-800/80">
          <h2 className="text-xs font-bold tracking-wider text-slate-300 uppercase">ANÁLISE DAS DATAS DE REMESSA</h2>
          <p className="text-[11px] text-slate-400 mb-4">Volume de itens por cumprimento do prazo de remessa prometido</p>
          <div className="h-64">
            {dadosFiltrados.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={analiseRemessaData} margin={{ top: 15, right: 30, left: 0, bottom: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1E293B" vertical={false} />
                  <XAxis dataKey="status" stroke="#64748B" fontSize={10} />
                  <YAxis stroke="#64748B" fontSize={10} allowDecimals={false} />
                  <Tooltip
                    contentStyle={{ backgroundColor: '#171E2E', borderColor: '#334155', borderRadius: '8px' }}
                    formatter={(value) => [`${value} itens`, 'Volume']}
                  />
                  <Bar dataKey="qtd" radius={[4, 4, 0, 0]} name="Itens">
                    {analiseRemessaData.map((entry, index) => (
                      <Cell key={`cell-remessa-${index}`} fill={entry.fill} />
                    ))}
                    <LabelList dataKey="qtd" position="top" fill="#94A3B8" fontSize={10} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-xs text-slate-500">
                Faça o upload do ficheiro Excel para carregar a análise de remessa.
              </div>
            )}
          </div>
        </div>

      </div>

    </div>
  );
}