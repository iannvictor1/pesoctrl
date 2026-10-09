import { useEffect, useState } from 'react';
import { ArrowLeft, ArrowUpRight, Boxes, ClipboardList, Download, Package, RefreshCw, Search, Smartphone, Weight } from 'lucide-react';
import { request, exportarInventario } from './api';

type Resumo = { id: number; operador_nome: string | null; status: string; inicio: string; fim: string | null; total_pallets: number; total_caixas: number; peso_total_kg: number | string };
type Caixa = { id: number; codigo: string; descricao: string; validade: string; peso_kg: number | string; coletor_nome: string | null; data_hora: string };
type Pallet = { unitizador: string | null; endereco: string | null; id: number; status: string; inicio: string; fim: string | null; caixas: Caixa[] };
type Detalhe = Omit<Resumo, 'total_pallets' | 'total_caixas' | 'peso_total_kg'> & { pallets: Pallet[] };
const kg = (n: number | string) => Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
const data = (s: string) => new Date(s).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });

export default function ConsultaInventarios() {
  const [lista, setLista] = useState<Resumo[]>([]);
  const [id, setId] = useState<number | null>(null);
  const [detalhe, setDetalhe] = useState<Detalhe | null>(null);
  const [busca, setBusca] = useState('');
  const [status, setStatus] = useState('todos');
  const [erro, setErro] = useState('');
  const [carregando, setCarregando] = useState(true);
  const [atualizacao, setAtualizacao] = useState('');
  const [versao, setVersao] = useState(0);
  const [exportando, setExportando] = useState(false);
  const [exportacao, setExportacao] = useState('');

  async function baixarPlanilha() {
    if (!id || exportando) return;
    setExportando(true); setExportacao('');
    try { await exportarInventario(id); setExportacao(`Download da planilha do inventário ${id} iniciado.`); }
    catch (e) { setExportacao(e instanceof Error ? e.message : 'Falha ao exportar a planilha.'); }
    finally { setExportando(false); }
  }

  useEffect(() => {
    const controller = new AbortController();
    let emCurso = false;
    setCarregando(true); setErro('');
    async function carregar() {
      if (emCurso) return;
      emCurso = true;
      try {
        const options = { signal: controller.signal };
        if (id) {
          const result = await request<Detalhe>(`/api/inventarios/${id}`, {}, options);
          if (!controller.signal.aborted) setDetalhe(result);
        } else {
          const result = await request<Resumo[]>('/api/inventarios/resumos', {}, options);
          if (!controller.signal.aborted) setLista(result);
        }
        if (!controller.signal.aborted) { setErro(''); setAtualizacao(new Date().toLocaleTimeString('pt-BR')); }
      } catch (e) {
        if (!controller.signal.aborted) setErro(e instanceof Error ? e.message : 'Não foi possível carregar os inventários.');
      } finally { emCurso = false; if (!controller.signal.aborted) setCarregando(false); }
    }
    carregar();
    const timer = window.setInterval(carregar, 10000);
    return () => { controller.abort(); window.clearInterval(timer); };
  }, [id, versao]);

  const filtrados = lista.filter(i => (status === 'todos' || i.status === status) && `inventário ${i.id}`.includes(busca.trim().toLowerCase()));
  const caixas = detalhe?.pallets.flatMap(p => p.caixas) ?? [];
  function abrir(numero: number | null) { setDetalhe(null); setId(numero); }
  return <section className="inventory inv-workspace inv-consult">
    <header className="inv-header"><div className="inv-title"><span className="inv-title-icon"><ClipboardList size={26} /></span><div><span className="inv-eyebrow">CONSULTA DE ESTOQUE</span><h2>{id ? `Inventário ${id}` : 'Consulta de inventários'}</h2></div></div><div className="inv-sync"><span>{atualizacao ? `Atualizado às ${atualizacao}` : 'Carregando…'}</span><button disabled={carregando} aria-label="Atualizar consulta" onClick={() => setVersao(n => n + 1)}><RefreshCw size={18} /></button></div></header>
    {id ? <div className="inv-export-toolbar"><button className="inv-back" disabled={exportando} onClick={() => { setExportacao(''); abrir(null); }}><ArrowLeft size={18} />Todos os inventários</button><button className="inv-create" disabled={!detalhe || exportando} onClick={baixarPlanilha}><Download size={18} />{exportando ? 'Gerando planilha…' : 'Exportar planilha (.xlsx)'}</button></div> : <p>Selecione um quadro para consultar os pallets e as caixas do inventário.</p>}
    {id && exportacao && <p role="status">{exportacao}</p>}
    {erro && <div className="alert" role="alert">{erro}</div>}
    {!id && <>
      <div className="inv-query-filters"><label><span><Search size={15} />Buscar inventário</span><input value={busca} onChange={e => setBusca(e.target.value)} placeholder="Ex.: 3" /></label><label>Status<select value={status} onChange={e => setStatus(e.target.value)}><option value="todos">Todos os status</option><option value="aberto">Abertos</option><option value="encerrado">Encerrados</option></select></label></div>
      <div className="inv-query-grid">{filtrados.map(i => <button className="inv-summary-card" key={i.id} onClick={() => abrir(i.id)} aria-label={`Abrir inventário ${i.id}`}>
        <div className="inv-summary-heading"><span className="inv-summary-icon"><ClipboardList size={23} /></span><span className={`inv-badge ${i.status}`}>{i.status === 'aberto' ? 'Aberto' : 'Encerrado'}</span></div>
        <h3>Inventário {i.id}</h3><span className="inv-card-date">Iniciado em {data(i.inicio)}</span>
        <span className="inv-card-date">Operador: {i.operador_nome || 'Não informado'}</span>
        <div className="inv-summary-counts"><span><Boxes size={17} /><strong>{i.total_pallets}</strong> pallets</span><span><Package size={17} /><strong>{i.total_caixas}</strong> caixas</span></div>
        <div className="inv-summary-weight"><Weight size={19} /><strong>{kg(i.peso_total_kg)} <small>kg</small></strong></div>
        <div className="inv-card-bottom"><span>Ver detalhes</span><ArrowUpRight size={18} /></div>
      </button>)}</div>
      {!filtrados.length && <div className="inv-card inv-empty"><ClipboardList size={32} /><strong>{carregando ? 'Carregando inventários…' : lista.length ? 'Nenhum inventário encontrado' : 'Nenhum inventário cadastrado'}</strong><p>{lista.length ? 'Experimente outro número ou status.' : 'Os inventários criados na aba Inventário aparecerão aqui.'}</p></div>}
    </>}
    {id && !detalhe && !erro && <div className="inv-empty">Carregando detalhes…</div>}
    {id && detalhe && <>
      <p>Operador: <strong>{detalhe.operador_nome || 'Não informado'}</strong></p>
      <div className="inv-detail-meta"><span className={`inv-badge ${detalhe.status}`}>{detalhe.status === 'aberto' ? 'Aberto' : 'Encerrado'}</span><span>Início: {data(detalhe.inicio)}</span>{detalhe.fim && <span>Encerramento: {data(detalhe.fim)}</span>}</div>
      <div className="inv-stats inv-query-stats"><article><Boxes size={22} /><div><span>Pallets</span><strong>{detalhe.pallets.length}</strong></div></article><article><Package size={22} /><div><span>Caixas</span><strong>{caixas.length}</strong></div></article><article><Weight size={22} /><div><span>Peso total</span><strong>{kg(caixas.reduce((s, c) => s + Number(c.peso_kg), 0))} <small>kg</small></strong></div></article></div>
      <div className="inv-section-heading"><Boxes size={20} /><h3>Pallets deste inventário</h3><span className="inv-caption">Abra um pallet para ver suas caixas</span></div>
      {!detalhe.pallets.length && <div className="inv-card inv-empty"><Boxes size={32} /><strong>Nenhum pallet criado neste inventário</strong></div>}
      {detalhe.pallets.map(p => <details className="inv-card inv-detail-pallet" key={p.id}><summary><div><strong>Pallet {p.id}</strong><span className={`inv-badge ${p.status}`}>{p.status === 'aberto' ? 'Aberto' : 'Encerrado'}</span></div><span>{p.caixas.length} caixas · {kg(p.caixas.reduce((s, c) => s + Number(c.peso_kg), 0))} kg</span></summary><p>Unitizador: <strong>{p.unitizador || 'Não informado'}</strong> · Endereço: <strong>{p.endereco || 'Não informado'}</strong></p><p className="inv-pallet-dates">Início: {data(p.inicio)}{p.fim ? ` · Encerramento: ${data(p.fim)}` : ''}</p>
        {!p.caixas.length ? <div className="inv-empty">Nenhuma caixa registrada neste pallet.</div> : <div className="inv-box-list">{p.caixas.map(c => <article className="inv-box" key={c.id}><div className="inv-box-main"><span className="inv-box-number">#{c.id}</span><h4>{c.descricao}</h4><strong>{kg(c.peso_kg)} <small>kg</small></strong></div><div className="inv-box-meta"><span>Validade <b>{c.validade}</b></span><span><Smartphone size={14} />{c.coletor_nome || 'Sem identificação'}</span><span>{data(c.data_hora)}</span></div><code>{c.codigo}</code></article>)}</div>}
      </details>)}
    </>}
    <footer className="inv-footer">Consulta atualizada automaticamente a cada 10 segundos.</footer>
  </section>;
}
