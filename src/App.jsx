import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LabelList
} from 'recharts';

export default function PainelProcurement() {
  const [dados, setDados] = useState([]);
  const [filtroGC, setFiltroGC] = useState('Todos');
  const [filtroCentro, setFiltroCentro] = useState('Todos');
  const [dataSnapshot, setDataSnapshot] = useState('Snapshot Atual');

  // --- LEITURA E TRATAMENTO DA PLANILHA DE PROCUREMENT ---
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const dataParsed = XLSX.utils.sheet_to_json(ws);

      const formatados = dataParsed.map((item, idx) => ({
        id: item['Concat'] || `${item['Requisição de compra']}-${item['Item ReqC']}` || idx,
        gcCompras: item['GC Compras'] || item['Comprador'] || 'Não Atribuído',
        acaoRC: item['Ação RC'] || 'Outros',
        obs: item['OBS'] || 'Sem OBS',
        tratativa: item['Tratativa'] ? String(item['Tratativa']).trim() : null,
        temPedido: Boolean(item['Tratativa'] && String(item['Tratativa']).trim() !== ''),
        centro: item['Nome Centro'] || item['Centro'] || 'Geral',
        diasRC: Number(item['Dias RC'] || 0),
        dataSolicitacao: item['Data da solicitação'] || '',
        dateSnapshot: item['SnapshotDate'] || new Date().toLocaleDateString('pt-BR')
      }));

      setDados(formatados);
    };
    reader.readAsBinaryString(file);
  };

  // --- FILTRAGEM DOS DADOS ---
  const dadosFiltrados = useMemo(() => {
    return dados.filter((item) => {
      const matchGC = filtroGC === 'Todos' || item.gcCompras === filtroGC;
      const matchCentro = filtroCentro === 'Todos' || item.centro === filtroCentro;
      return matchGC && matchCentro;
    });
  }, [dados, filtroGC, filtroCentro]);

  // Total Geral Filtrado
  const totalGeral = dadosFiltrados.length || 1;

  // --- 1. QUADRANTE EVOLUÇÃO TEMPORAL (ESTILO BI DO LAURENCE) ---
  const evolucaoPorGC = useMemo(() => {
    const agrupado = {};

    dadosFiltrados.forEach((item) => {
      const gc = item.gcCompras;
      if (!agrupado[gc]) agrupado[gc] = { gc, total: 0 };
      agrupado[gc].total += 1;
    });

    return Object.values(agrupado)
      .map(d => ({
        ...d,
        porcentagem: ((d.total / totalGeral) * 100).toFixed(1)
      }))
      .sort((a, b) => b.total - a.total);
  }, [dadosFiltrados, totalGeral]);

  // --- 2. QUADRANTE EVOLUÇÃO DAS TRATATIVAS (GERARAM PEDIDO VS NÃO GERARAM) ---
  const evolucaoTratativas = useMemo(() => {
    const geraramPedido = dadosFiltrados.filter(d => d.temPedido).length;
    const naoGeraramPedido = dadosFiltrados.length - geraramPedido;

    const pctGeraram = ((geraramPedido / totalGeral) * 100).toFixed(1);
    const pctNaoGeraram = ((naoGeraramPedido / totalGeral) * 100).toFixed(1);

    return [
      { name: 'Geraram Pedido', valor: geraramPedido, pct: pctGeraram, fill: '#00a86b' },
      { name: 'Não Geraram Pedido', valor: naoGeraramPedido, pct: pctNaoGeraram, fill: '#e74c3c' }
    ];
  }, [dadosFiltrados, totalGeral]);

  // --- 3. QUADRANTE PRINCIPAIS MOTIVOS DE PENDÊNCIAS (OBS / AÇÃO RC) ---
  const motivosPendencias = useMemo(() => {
    const pendentes = dadosFiltrados.filter(d => !d.temPedido);
    const totalPendentes = pendentes.length || 1;
    const contagem = {};

    pendentes.forEach(d => {
      const motivo = d.obs !== 'Sem OBS' ? d.obs : (d.acaoRC || 'Sem Motivo Especificado');
      contagem[motivo] = (contagem[motivo] || 0) + 1;
    });

    return Object.keys(contagem)
      .map(m => ({
        motivo: m,
        qtd: contagem[m],
        pct: ((contagem[m] / totalPendentes) * 100).toFixed(1),
        pctGeral: ((contagem[m] / totalGeral) * 100).toFixed(1)
      }))
      .sort((a, b) => b.qtd - a.qtd)
      .slice(0, 8); // Top 8 motivos
  }, [dadosFiltrados, totalGeral]);

  // Listas para Filtros
  const listaGC = useMemo(() => ['Todos', ...new Set(dados.map(d => d.gcCompras))], [dados]);
  const listaCentros = useMemo(() => ['Todos', ...new Set(dados.map(d => d.centro))], [dados]);

  return (
    <div style={{ backgroundColor: '#f4f6f9', minHeight: '100vh', padding: '20px', fontFamily: 'Segoe UI, sans-serif' }}>
      
      {/* CABEÇALHO */}
      <div style={{ backgroundColor: '#002060', color: '#fff', padding: '15px 20px', borderRadius: '8px', marginBottom: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px' }}>
        <div>
          <h2 style={{ margin: 0, fontSize: '20px' }}>Dashboard de Requisições Pendentes</h2>
          <span style={{ fontSize: '12px', opacity: 0.8 }}>Visualização com Valores Quantitativos e Percentuais (%)</span>
        </div>
        
        <div style={{ display: 'flex', gap: '15px', alignItems: 'center', flexWrap: 'wrap' }}>
          <div>
            <label style={{ fontSize: '11px', display: 'block' }}>Carregar Excel (.xlsx):</label>
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} style={{ color: '#fff', fontSize: '12px' }} />
          </div>
          <div>
            <label style={{ fontSize: '11px', display: 'block' }}>Gestão Compras (GC):</label>
            <select value={filtroGC} onChange={e => setFiltroGC(e.target.value)} style={{ padding: '5px', borderRadius: '4px' }}>
              {listaGC.map(gc => <option key={gc} value={gc}>{gc}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: '11px', display: 'block' }}>Centro:</label>
            <select value={filtroCentro} onChange={e => setFiltroCentro(e.target.value)} style={{ padding: '5px', borderRadius: '4px' }}>
              {listaCentros.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>
      </div>

      {/* CARDS DE RESUMO (QUANTITATIVOS E PERCENTUAIS) */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '15px', marginBottom: '20px' }}>
        <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', borderLeft: '5px solid #002060', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <span style={{ fontSize: '12px', color: '#666' }}>Total de Requisições</span>
          <h3 style={{ margin: '5px 0 0 0', color: '#002060' }}>{dadosFiltrados.length} <span style={{ fontSize: '14px', color: '#666' }}>(100%)</span></h3>
        </div>
        <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', borderLeft: '5px solid #00a86b', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <span style={{ fontSize: '12px', color: '#666' }}>Geraram Pedido</span>
          <h3 style={{ margin: '5px 0 0 0', color: '#00a86b' }}>
            {evolucaoTratativas[0]?.valor} <span style={{ fontSize: '14px' }}>({evolucaoTratativas[0]?.pct}%)</span>
          </h3>
        </div>
        <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', borderLeft: '5px solid #e74c3c', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <span style={{ fontSize: '12px', color: '#666' }}>Não Geraram Pedido (Pendentes)</span>
          <h3 style={{ margin: '5px 0 0 0', color: '#e74c3c' }}>
            {evolucaoTratativas[1]?.valor} <span style={{ fontSize: '14px' }}>({evolucaoTratativas[1]?.pct}%)</span>
          </h3>
        </div>
      </div>

      {/* GRID DE QUADRANTES */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '20px' }}>
        
        {/* QUADRANTE 1: EVOLUÇÃO TEMPORAL (ESTILO BI LAURENCE BY GC COMPRAS) */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', gridColumn: 'span 2' }}>
          <h4 style={{ margin: '0 0 10px 0', color: '#002060' }}>Evolução Temporal / Qtd. Requisições por Gestão de Compras</h4>
          <div style={{ width: '100%', height: 320 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={evolucaoPorGC} margin={{ top: 20, right: 30, left: 0, bottom: 50 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} />
                <XAxis dataKey="gc" interval={0} angle={-25} textAnchor="end" />
                <YAxis />
                <Tooltip formatter={(value, name, item) => [`${value} un (${item.payload.porcentagem}%)`, 'Volume']} />
                <Bar dataKey="total" fill="#002060">
                  <LabelList dataKey="total" position="top" formatter={(val) => val} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* QUADRANTE 2: EVOLUÇÃO DAS TRATATIVAS (PEDIDOS GERADOS VS NÃO GERADOS) */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <h4 style={{ margin: '0 0 15px 0', color: '#002060' }}>Evolução das Tratativas (Resultado RCs)</h4>
          <div style={{ width: '100%', height: 260 }}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={evolucaoTratativas} dataKey="valor" nameKey="name" cx="50%" cy="50%" outerRadius={80} label={({ name, valor, pct }) => `${name}: ${valor} (${pct}%)`}>
                  {evolucaoTratativas.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.fill} />
                  ))}
                </Pie>
                <Tooltip formatter={(value, name, props) => [`${value} un (${props.payload.pct}%)`, name]} />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* QUADRANTE 3: PRINCIPAIS MOTIVOS DE PENDÊNCIAS */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <h4 style={{ margin: '0 0 15px 0', color: '#002060' }}>Principais Motivos de Pendências (OBS / Ação)</h4>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
              <thead>
                <tr style={{ backgroundColor: '#f0f4f8', borderBottom: '2px solid #ddd', textAlign: 'left' }}>
                  <th style={{ padding: '8px' }}>Motivo / Observação</th>
                  <th style={{ padding: '8px', textAlign: 'right' }}>Qtd. Absoluta</th>
                  <th style={{ padding: '8px', textAlign: 'right' }}>% do Total Pendente</th>
                </tr>
              </thead>
              <tbody>
                {motivosPendencias.map((m, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid #eee' }}>
                    <td style={{ padding: '8px' }}>{m.motivo}</td>
                    <td style={{ padding: '8px', textAlign: 'right', fontWeight: 'bold' }}>{m.qtd}</td>
                    <td style={{ padding: '8px', textAlign: 'right', color: '#e74c3c', fontWeight: 'bold' }}>{m.pct}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}