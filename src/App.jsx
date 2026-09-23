import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  BarChart, Bar, PieChart, Pie, Cell, LabelList, ComposedChart, Line, Legend
} from 'recharts';

// ------------------------------------------------------------------
// HELPERS E UTILITÁRIOS
// ------------------------------------------------------------------
const extrairComprador = (acao) => {
  if (!acao) return 'Não atribuído';
  const acaoStr = String(acao);
  return acaoStr.includes(':') ? acaoStr.split(':')[1].trim() : acaoStr;
};

const calcularFaixaAging = (dias) => {
  const d = Number(dias) || 0;
  if (d <= 15) return '0 - 15 dias';
  if (d <= 30) return '16 - 30 dias';
  if (d <= 60) return '31 - 60 dias';
  return 'Acima de 60 dias';
};

const normalizarObs = (obs) => {
  if (!obs) return 'SEM MOTIVO DECLARADO';
  return String(obs).trim().replace(/\s+/g, ' ').toUpperCase();
};

const parseData = (str) => {
  if (!str) return null;
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
};

const formatarData = (str) => {
  const d = parseData(str);
  return d ? d.toLocaleDateString('pt-BR') : '—';
};

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884d8', '#82ca9d'];

// ------------------------------------------------------------------
// COMPONENTE PRINCIPAL DO DASHBOARD
// ------------------------------------------------------------------
export default function DashboardProcurement() {
  const [dados, setDados] = useState([]);
  const [filtroComprador, setFiltroComprador] = useState('TODOS');
  const [filtroCentro, setFiltroCentro] = useState('TODOS');

  // Manipulador de Upload de Planilha (.xlsx, .xls, .csv)
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const dataRaw = XLSX.utils.sheet_to_json(ws);

      // Mapeamento dinâmico conforme estrutura da planilha inserida
      const dadosTratados = dataRaw.map((row, idx) => {
        const diasRC = Number(row['Dias RC'] || row['diasRC'] || 0);
        return {
          id: row['ID'] || idx,
          rc: String(row['Requisição de compras'] || row['rc'] || ''),
          itemRc: row['Item da requisição de compras'] || row['itemRc'],
          diasRC: diasRC,
          centro: row['Centro'] || row['centro'] || 'Não informado',
          comprador: extrairComprador(row['Última ação efetuada por'] || row['comprador']),
          grupoCompras: row['Grupo de compras'] || row['grupoCompras'],
          faixaAging: calcularFaixaAging(diasRC),
          qtdSolicitada: Number(row['Quantidade solicitada'] || row['qtdSolicitada'] || 0),
          material: String(row['Material'] || row['material'] || ''),
          textoBreve: row['Texto breve'] || row['textoBreve'] || '',
          dataSolicitacao: row['Data da solicitação'] || row['dataSolicitacao'],
          remessa: row['Data de remessa'] || row['remessa'],
          tratativa: row['Tratativa'] || row['tratativa'] || null,
          obs: row['Observação do Comprador'] || row['obs'] || null
        };
      });

      setDados(dadosTratados);
    };
    reader.readAsBinaryString(file);
  };

  // Filtragem dos dados
  const dadosFiltrados = useMemo(() => {
    return dados.filter(item => {
      const matchComprador = filtroComprador === 'TODOS' || item.comprador === filtroComprador;
      const matchCentro = filtroCentro === 'TODOS' || item.centro === filtroCentro;
      return matchComprador && matchCentro;
    });
  }, [dados, filtroComprador, filtroCentro]);

  // Lista de Filtros Únicos
  const compradoresUnicos = useMemo(() => Array.from(new Set(dados.map(d => d.comprador))), [dados]);
  const centrosUnicos = useMemo(() => Array.from(new Set(dados.map(d => d.centro))), [dados]);

  // 1. MÉTRICAS E KPIs
  const totalItens = dadosFiltrados.length;
  const totalComPedido = useMemo(() => {
    return dadosFiltrados.filter(d => {
      const t = String(d.tratativa || '').toLowerCase();
      const o = String(d.obs || '').toLowerCase();
      return t.includes('gerar pedido') || t.includes('pedido gerado') || o.includes('pedido gerado');
    }).length;
  }, [dadosFiltrados]);

  const totalSemPedido = totalItens - totalComPedido;
  const percComPedido = totalItens ? ((totalComPedido / totalItens) * 100).toFixed(1) : '0.0';
  const percSemPedido = totalItens ? ((totalSemPedido / totalItens) * 100).toFixed(1) : '0.0';

  // 2. EVOLUÇÃO TEMPORAL (Padrão BI Laurence: ComposedChart - Barras + Linha de Tendência)
  const dadosEvolucaoTemporal = useMemo(() => {
    const agrupa = {};
    dadosFiltrados.forEach(item => {
      const data = item.dataSolicitacao ? formatarData(item.dataSolicitacao) : 'Sem Data';
      if (!agrupa[data]) agrupa[data] = 0;
      agrupa[data] += 1;
    });

    let acumulado = 0;
    return Object.keys(agrupa).map(data => {
      acumulado += agrupa[data];
      const pct = totalItens ? ((agrupa[data] / totalItens) * 100).toFixed(1) : 0;
      return {
        data,
        qtd: agrupa[data],
        pct: Number(pct),
        qtdAcumulada: acumulado
      };
    });
  }, [dadosFiltrados, totalItens]);

  // 3. MOTIVOS DE PENDÊNCIAS E ANÁLISE DE ATRASO DA REMESSA (Integrado)
  const dadosMotivosAtraso = useMemo(() => {
    const agrupa = {};
    const hoje = new Date();

    dadosFiltrados.forEach(item => {
      let motivo = normalizarObs(item.obs);
      const dataRemessa = parseData(item.remessa);
      
      // Validação de atraso na data de remessa
      if (dataRemessa && dataRemessa < hoje) {
        motivo = `[FORA DO PRAZO REMESSA] ${motivo}`;
      }

      if (!agrupa[motivo]) agrupa[motivo] = 0;
      agrupa[motivo] += 1;
    });

    return Object.keys(agrupa).map(motivo => ({
      motivo,
      qtd: agrupa[motivo],
      pct: totalItens ? ((agrupa[motivo] / totalItens) * 100).toFixed(1) : '0.0'
    })).sort((a, b) => b.qtd - a.qtd);
  }, [dadosFiltrados, totalItens]);

  // 4. EVOLUÇÃO DAS TRATATIVAS (Lógica Solicitada: Entraram x Geraram Pedido x Não Geraram)
  const dadosTratativasFunil = useMemo(() => {
    return [
      { name: 'Com Pedido Gerado', valor: totalComPedido, pct: percComPedido },
      { name: 'Sem Pedido / Em Tratativa', valor: totalSemPedido, pct: percSemPedido }
    ];
  }, [totalComPedido, totalSemPedido, percComPedido, percSemPedido]);

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif', backgroundColor: '#f4f6f9' }}>
      <h2>Dashboard de Procurements & Requisições</h2>

      {/* ÁREA DE UPLOAD E FILTROS */}
      <div style={{ display: 'flex', gap: '20px', marginBottom: '20px', background: '#fff', padding: '15px', borderRadius: '8px' }}>
        <div>
          <label><strong>Upload Planilha (.xlsx): </strong></label>
          <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} />
        </div>
        <div>
          <label><strong>Comprador: </strong></label>
          <select value={filtroComprador} onChange={e => setFiltroComprador(e.target.value)}>
            <option value="TODOS">Todos</option>
            {compradoresUnicos.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <label><strong>Centro: </strong></label>
          <select value={filtroCentro} onChange={e => setFiltroCentro(e.target.value)}>
            <option value="TODOS">Todos</option>
            {centrosUnicos.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
      </div>

      {/* KPI CARDS (DADOS QUANTITATIVOS E PERCENTUAIS) */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '15px', marginBottom: '20px' }}>
        <div style={{ background: '#fff', padding: '15px', borderRadius: '8px', borderLeft: '5px solid #0088FE' }}>
          <h4>Total de RCs Analisadas</h4>
          <h3>{totalItens} <span style={{ fontSize: '14px', color: '#666' }}>(100%)</span></h3>
        </div>
        <div style={{ background: '#fff', padding: '15px', borderRadius: '8px', borderLeft: '5px solid #00C49F' }}>
          <h4>RCs com Pedido Gerado</h4>
          <h3>{totalComPedido} <span style={{ fontSize: '14px', color: '#666' }}>({percComPedido}%)</span></h3>
        </div>
        <div style={{ background: '#fff', padding: '15px', borderRadius: '8px', borderLeft: '5px solid #FF8042' }}>
          <h4>RCs Pendentes / Sem Pedido</h4>
          <h3>{totalSemPedido} <span style={{ fontSize: '14px', color: '#666' }}>({percSemPedido}%)</span></h3>
        </div>
      </div>

      {/* QUADRANTES VISUAIS */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
        
        {/* QUADRANTE 1: EVOLUÇÃO TEMPORAL (BI LAURENCE - COMPOSED CHART) */}
        <div style={{ background: '#fff', padding: '15px', borderRadius: '8px' }}>
          <h3>Evolução Temporal das Solicitações (Padrão BI)</h3>
          <ResponsiveContainer width="100%" height={300}>
            <ComposedChart data={dadosEvolucaoTemporal}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="data" />
              <YAxis yAxisId="left" label={{ value: 'Qtd RCs', angle: -90, position: 'insideLeft' }} />
              <YAxis yAxisId="right" orientation="right" label={{ value: '% Vol', angle: 90, position: 'insideRight' }} />
              <Tooltip formatter={(value, name) => [name === 'pct' ? `${value}%` : value, name === 'qtd' ? 'Volume Qtd' : 'Proporção %']} />
              <Legend />
              <Bar yAxisId="left" dataKey="qtd" name="Volume RCs (Qtd)" fill="#0088FE" />
              <Line yAxisId="right" type="monotone" dataKey="pct" name="Participação (%)" stroke="#FF8042" strokeWidth={2} />
            </ComposedChart>
          </ResponsiveContainer>
        </div>

        {/* QUADRANTE 2: EVOLUÇÃO DAS TRATATIVAS (LÓGICA CONVERSÃO/FUNIL) */}
        <div style={{ background: '#fff', padding: '15px', borderRadius: '8px' }}>
          <h3>Evolução das Tratativas (Resultado RCs)</h3>
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={dadosTratativasFunil}
                cx="50%"
                cy="50%"
                outerRadius={100}
                dataKey="valor"
                label={({ name, valor, pct }) => `${name}: ${valor} (${pct}%)`}
              >
                {dadosTratativasFunil.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip formatter={(value) => [`${value} RCs`, 'Quantidade']} />
            </PieChart>
          </ResponsiveContainer>
        </div>

        {/* QUADRANTE 3: PRINCIPAIS MOTIVOS DE PENDÊNCIAS E ANÁLISE DE ATRASO NA REMESSA */}
        <div style={{ background: '#fff', padding: '15px', borderRadius: '8px', gridColumn: 'span 2' }}>
          <h3>Principais Motivos de Pendências e Cumprimento da Data de Remessa</h3>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={dadosMotivosAtraso} layout="vertical" margin={{ left: 150 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis type="number" />
              <YAxis type="category" dataKey="motivo" width={220} style={{ fontSize: '11px' }} />
              <Tooltip formatter={(value, name, props) => [`${value} RCs (${props.payload.pct}%)`, 'Total']} />
              <Bar dataKey="qtd" fill="#8884d8">
                <LabelList dataKey="qtd" position="right" formatter={(val, entry) => `${val} un. (${dadosMotivosAtraso.find(d => d.qtd === val)?.pct || 0}%)`} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

      </div>
    </div>
  );
}