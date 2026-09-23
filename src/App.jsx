import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ComposedChart,
  Bar,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer
} from 'recharts';

// Dados de exemplo (substituídos se fizer upload do Excel)
const dadosIniciais = [
  { id: 1, dataSolicitacao: '2026-01-10', comprador: 'Ana Silva', centro: 'Planta A', tratativa: true, diasRC: 5 },
  { id: 2, dataSolicitacao: '2026-01-15', comprador: 'Carlos Souza', centro: 'Planta B', tratativa: false, diasRC: 12 },
  { id: 3, dataSolicitacao: '2026-02-01', comprador: 'Ana Silva', centro: 'Planta B', tratativa: true, diasRC: 8 },
  { id: 4, dataSolicitacao: '2026-02-18', comprador: 'Mariana Lima', centro: 'Planta A', tratativa: true, diasRC: 4 },
  { id: 5, dataSolicitacao: '2026-03-05', comprador: 'Carlos Souza', centro: 'Planta A', tratativa: false, diasRC: 20 },
  { id: 6, dataSolicitacao: '2026-03-20', comprador: 'Ana Silva', centro: 'Planta B', tratativa: true, diasRC: 3 }
];

export default function PainelProcurement() {
  const [dados, setDados] = useState(dadosIniciais);
  const [abaAtiva, setAbaAtiva] = useState('evolucao');
  const [filtroComprador, setFiltroComprador] = useState('Todos');
  const [filtroCentro, setFiltroCentro] = useState('Todos');

  // --- LEITURA DO ARQUIVO EXCEL ---
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
        id: item.id || idx + 1,
        dataSolicitacao: item.dataSolicitacao || item.Data || '',
        comprador: item.comprador || item.Comprador || 'Não atribuído',
        centro: item.centro || item.Centro || 'Geral',
        tratativa: item.tratativa !== undefined ? Boolean(item.tratativa) : Boolean(item.Status === 'Concluído'),
        diasRC: Number(item.diasRC || item.Dias || 0)
      }));

      setDados(formatados);
    };
    reader.readAsBinaryString(file);
  };

  // --- FILTRAGEM ---
  const dadosFiltrados = useMemo(() => {
    return dados.filter((item) => {
      const matchComprador = filtroComprador === 'Todos' || item.comprador === filtroComprador;
      const matchCentro = filtroCentro === 'Todos' || item.centro === filtroCentro;
      return matchComprador && matchCentro;
    });
  }, [dados, filtroComprador, filtroCentro]);

  // --- AGREGADOR MENSAL PARA A EVOLUÇÃO TEMPORAL ---
  const evolucaoTemporal = useMemo(() => {
    const agrupado = {};

    dadosFiltrados.forEach((item) => {
      if (!item.dataSolicitacao) return;

      let anoMes = '';
      if (item.dataSolicitacao.includes('-')) {
        anoMes = item.dataSolicitacao.substring(0, 7);
      } else if (item.dataSolicitacao.includes('/')) {
        const partes = item.dataSolicitacao.split('/');
        anoMes = `${partes[2]}-${partes[1].padStart(2, '0')}`;
      } else {
        anoMes = 'Outros';
      }

      if (!agrupado[anoMes]) {
        agrupado[anoMes] = {
          mes: anoMes,
          totalRCs: 0,
          comTratativa: 0,
          semTratativa: 0,
          somaDias: 0,
          agingMedio: 0
        };
      }

      agrupado[anoMes].totalRCs += 1;
      agrupado[anoMes].somaDias += Number(item.diasRC) || 0;

      if (item.tratativa) {
        agrupado[anoMes].comTratativa += 1;
      } else {
        agrupado[anoMes].semTratativa += 1;
      }
    });

    return Object.values(agrupado)
      .map((d) => ({
        ...d,
        agingMedio: d.totalRCs > 0 ? Math.round(d.somaDias / d.totalRCs) : 0
      }))
      .sort((a, b) => a.mes.localeCompare(b.mes));
  }, [dadosFiltrados]);

  const compradores = useMemo(() => ['Todos', ...new Set(dados.map((d) => d.comprador))], [dados]);
  const centros = useMemo(() => ['Todos', ...new Set(dados.map((d) => d.centro))], [dados]);

  return (
    <div style={{ backgroundColor: '#f4f6f9', minHeight: '100vh', padding: '20px', fontFamily: 'Segoe UI, sans-serif' }}>
      {/* BARRA SUPERIOR DE FILTROS E IMPORTAÇÃO */}
      <div style={{ backgroundColor: '#002060', color: '#fff', padding: '15px 20px', borderRadius: '8px', marginBottom: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px' }}>
        <h2 style={{ margin: 0, fontSize: '20px' }}>Dashboard de Procurement — Requisições de Compras</h2>
        
        <div style={{ display: 'flex', gap: '15px', alignItems: 'center', flexWrap: 'wrap' }}>
          <div>
            <label style={{ fontSize: '12px', display: 'block', marginBottom: '2px' }}>Carregar Excel:</label>
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} style={{ color: '#fff', fontSize: '12px' }} />
          </div>
          <div>
            <label style={{ fontSize: '12px', display: 'block', marginBottom: '2px' }}>Comprador:</label>
            <select value={filtroComprador} onChange={(e) => setFiltroComprador(e.target.value)} style={{ padding: '5px 10px', borderRadius: '4px', border: 'none' }}>
              {compradores.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>
          <div>
            <label style={{ fontSize: '12px', display: 'block', marginBottom: '2px' }}>Centro / Planta:</label>
            <select value={filtroCentro} onChange={(e) => setFiltroCentro(e.target.value)} style={{ padding: '5px 10px', borderRadius: '4px', border: 'none' }}>
              {centros.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* SELEÇÃO DE ABAS */}
      <div style={{ display: 'flex', gap: '10px', marginBottom: '20px' }}>
        <button
          onClick={() => setAbaAtiva('evolucao')}
          style={{ padding: '10px 20px', borderRadius: '6px', border: 'none', cursor: 'pointer', fontWeight: 'bold', backgroundColor: abaAtiva === 'evolucao' ? '#002060' : '#fff', color: abaAtiva === 'evolucao' ? '#fff' : '#333' }}
        >
          📈 Evolução Temporal (BI)
        </button>
        <button
          onClick={() => setAbaAtiva('detalhes')}
          style={{ padding: '10px 20px', borderRadius: '6px', border: 'none', cursor: 'pointer', fontWeight: 'bold', backgroundColor: abaAtiva === 'detalhes' ? '#002060' : '#fff', color: abaAtiva === 'detalhes' ? '#fff' : '#333' }}
        >
          📋 Base de Dados
        </button>
      </div>

      {/* ABA PRINCIPAL: EVOLUÇÃO TEMPORAL */}
      {abaAtiva === 'evolucao' && (
        <div>
          {/* PAINEL DE KPIS */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '15px', marginBottom: '20px' }}>
            <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', borderLeft: '5px solid #002060' }}>
              <span style={{ fontSize: '12px', color: '#666' }}>Total de RCs Solicitadas</span>
              <h3 style={{ margin: '5px 0 0 0', fontSize: '24px', color: '#002060' }}>{dadosFiltrados.length}</h3>
            </div>
            <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', borderLeft: '5px solid #00a86b' }}>
              <span style={{ fontSize: '12px', color: '#666' }}>RCs com Tratativa</span>
              <h3 style={{ margin: '5px 0 0 0', fontSize: '24px', color: '#00a86b' }}>
                {dadosFiltrados.filter((d) => d.tratativa).length}
              </h3>
            </div>
            <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', borderLeft: '5px solid #e74c3c' }}>
              <span style={{ fontSize: '12px', color: '#666' }}>Pendentes / Sem Tratativa</span>
              <h3 style={{ margin: '5px 0 0 0', fontSize: '24px', color: '#e74c3c' }}>
                {dadosFiltrados.filter((d) => !d.tratativa).length}
              </h3>
            </div>
            <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', borderLeft: '5px solid #f39c12' }}>
              <span style={{ fontSize: '12px', color: '#666' }}>Aging Médio Geral</span>
              <h3 style={{ margin: '5px 0 0 0', fontSize: '24px', color: '#f39c12' }}>
                {Math.round(
                  dadosFiltrados.reduce((acc, curr) => acc + (Number(curr.diasRC) || 0), 0) / (dadosFiltrados.length || 1)
                )}{' '}
                dias
              </h3>
            </div>
          </div>

          {/* ÁREA DA EVOLUÇÃO TEMPORAL (GRÁFICO MISTO) */}
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
            <h4 style={{ margin: '0 0 15px 0', color: '#002060' }}>Evolução Mensal de Solicitações e Aging Médio</h4>
            <div style={{ width: '100%', height: 380 }}>
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={evolucaoTemporal}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} />
                  <XAxis dataKey="mes" />
                  <YAxis yAxisId="left" orientation="left" stroke="#002060" />
                  <YAxis yAxisId="right" orientation="right" stroke="#f39c12" unit=" d" />
                  <Tooltip />
                  <Legend />
                  <Bar yAxisId="left" dataKey="comTratativa" name="Com Tratativa" stackId="a" fill="#00a86b" />
                  <Bar yAxisId="left" dataKey="semTratativa" name="Sem Tratativa" stackId="a" fill="#e74c3c" />
                  <Line yAxisId="right" type="monotone" dataKey="agingMedio" name="Aging Médio (dias)" stroke="#f39c12" strokeWidth={3} dot={{ r: 5 }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      )}

      {/* ABA: BASE DE DADOS */}
      {abaAtiva === 'detalhes' && (
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', overflowX: 'auto' }}>
          <h4 style={{ margin: '0 0 15px 0', color: '#002060' }}>Registros Filtrados ({dadosFiltrados.length})</h4>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '14px' }}>
            <thead>
              <tr style={{ backgroundColor: '#f0f4f8', borderBottom: '2px solid #ddd' }}>
                <th style={{ padding: '10px' }}>ID</th>
                <th style={{ padding: '10px' }}>Data Solicitação</th>
                <th style={{ padding: '10px' }}>Comprador</th>
                <th style={{ padding: '10px' }}>Centro</th>
                <th style={{ padding: '10px' }}>Tratativa</th>
                <th style={{ padding: '10px' }}>Dias RC</th>
              </tr>
            </thead>
            <tbody>
              {dadosFiltrados.map((item, index) => (
                <tr key={item.id || index} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '10px' }}>{item.id}</td>
                  <td style={{ padding: '10px' }}>{item.dataSolicitacao}</td>
                  <td style={{ padding: '10px' }}>{item.comprador}</td>
                  <td style={{ padding: '10px' }}>{item.centro}</td>
                  <td style={{ padding: '10px', fontWeight: 'bold', color: item.tratativa ? '#00a86b' : '#e74c3c' }}>
                    {item.tratativa ? 'Sim' : 'Não'}
                  </td>
                  <td style={{ padding: '10px' }}>{item.diasRC}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}