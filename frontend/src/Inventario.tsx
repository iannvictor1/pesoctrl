import { useEffect, useRef, useState } from 'react';
import { request } from './api';
import { Barcode, Boxes, CheckCircle2, ChevronRight, ClipboardList, Package, Plus, RefreshCw, Smartphone, Square, Weight } from 'lucide-react';

type Caixa = { coletor_nome?: string; id: number; codigo: string; descricao: string; validade: string; peso_kg: number | string };
type Pallet = { id: number; status: string; unitizador: string | null; endereco: string | null; endereco_codigo: string | null; caixas: Caixa[] };
type Inventario = { id: number; operador_nome: string | null; status: string; pallets: Pallet[] };
const kg = (n: number) => n.toLocaleString('pt-BR', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
const post = <T,>(path: string, body = {}) => request<T>(`/api/inventarios${path}`, {}, {
  method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
});

export default function InventarioTab() {
  const [coletorNome, setColetorNome] = useState(() => localStorage.getItem('inventario-coletor') || `Coletor ${Array.from(crypto.getRandomValues(new Uint8Array(3)), n => n.toString(16).padStart(2, '0')).join('').toUpperCase()}`);
  const selecao = useRef<{ inventario: number | null; pallet: number | null }>(JSON.parse(sessionStorage.getItem('inventario-selecao') || 'null') || { inventario: null, pallet: null });
  const revisao = useRef(0);
  const requisicao = useRef(0);
  const operando = useRef(false);
  const sincronizando = useRef(false);
  const [sincronizacao, setSincronizacao] = useState('');
  const [lista, setLista] = useState<Inventario[]>([]);
  const [operadorNome, setOperadorNome] = useState('');
  const [atual, setAtual] = useState<Inventario | null>(null);
  const [palletId, setPalletId] = useState<number | null>(null);
  const [codigo, setCodigo] = useState('');
  const [unitizador, setUnitizador] = useState('');
  const [enderecoCodigo, setEnderecoCodigo] = useState('');
  const unitizadorInput = useRef<HTMLInputElement>(null);
  const enderecoInput = useRef<HTMLInputElement>(null);
  const leituraRevisada = useRef<string | null>(null);
  const [dados, setDados] = useState({ descricao: '', validade: '', peso_kg: '' });
  const [pendentes, setPendentes] = useState(0);
  const [ocupado, setOcupado] = useState(false);
  const [erro, setErro] = useState('');
  const [mensagem, setMensagem] = useState('');
  const [falhas, setFalhas] = useState<{ codigo: string; mensagem: string; leituraId: string; inventarioId: number; palletId: number }[]>([]);
  const input = useRef<HTMLInputElement>(null);
  const fila = useRef(Promise.resolve());
  const pallet = atual?.pallets.find(p => p.id === palletId);
  const palletAberto = atual?.status === 'aberto' && pallet?.status === 'aberto';
  const habilitado = palletAberto && !!pallet?.unitizador && !!pallet?.endereco_codigo;
  const bloqueado = ocupado || pendentes > 0;

  function selecionarPallet(id: number | null) {
    if (selecao.current.pallet !== id) { setCodigo(''); setDados({ descricao: '', validade: '', peso_kg: '' }); setEnderecoCodigo(''); leituraRevisada.current = null; }
    selecao.current.pallet = id;
    sessionStorage.setItem('inventario-selecao', JSON.stringify(selecao.current));
    setPalletId(id);
  }
  async function carregar(id?: number, automatico = false) {
    const numero = ++requisicao.current;
    const versao = revisao.current;
    const items = await request<Inventario[]>('/api/inventarios');
    const selecionado = id ?? selecao.current.inventario;
    const detalhe = selecionado ? await request<Inventario>(`/api/inventarios/${selecionado}`) : null;
    if (numero !== requisicao.current || (automatico && (versao !== revisao.current || operando.current))) return;
    setLista(items);
    setAtual(detalhe);
    const anterior = selecao.current;
    selecao.current = { inventario: selecionado ?? null, pallet: anterior.inventario === selecionado && detalhe?.pallets.some(p => p.id === anterior.pallet) ? anterior.pallet : null };
    selecionarPallet(selecao.current.pallet);
    setSincronizacao(`Atualizado às ${new Date().toLocaleTimeString('pt-BR')}`);
  }
  useEffect(() => { localStorage.setItem('inventario-coletor', coletorNome); }, [coletorNome]);
  useEffect(() => {
    carregar().catch(e => setErro(e.message));
    const timer = window.setInterval(async () => {
      if (operando.current || sincronizando.current) return;
      sincronizando.current = true;
      try { await carregar(undefined, true); }
      catch { setSincronizacao('Sem atualização: confira a conexão com o servidor.'); }
      finally { sincronizando.current = false; }
    }, 3000);
    return () => { window.clearInterval(timer); requisicao.current++; };
  }, []);
  useEffect(() => {
    const proteger = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ''; };
    if (pendentes) window.addEventListener('beforeunload', proteger);
    return () => window.removeEventListener('beforeunload', proteger);
  }, [pendentes]);

  async function acao(fn: () => Promise<void>) {
    if (bloqueado || operando.current) return;
    operando.current = true; revisao.current++;
    setOcupado(true); setErro(''); setMensagem('');
    try { await fn(); } catch (e) { setErro(e instanceof Error ? e.message : 'Falha na operação.'); }
    finally { operando.current = false; revisao.current++; setOcupado(false); }
  }

  useEffect(() => {
    if (ocupado || atual?.status !== 'aberto') return;
    if (!pallet || pallet.status !== 'aberto') unitizadorInput.current?.focus();
    else if (!pallet.endereco_codigo) enderecoInput.current?.focus();
    else input.current?.focus();
  }, [atual?.id, palletId, pallet?.endereco_codigo, pallet?.status, ocupado]);

  function biparUnitizador() {
    if (!unitizador.trim() || atual?.status !== 'aberto') return;
    void acao(async () => {
      const novo = await post<Pallet>(`/${atual.id}/pallets`, { unitizador: unitizador.trim() });
      selecionarPallet(novo.id);
      setAtual(prev => prev?.id === atual.id ? { ...prev, pallets: [...prev.pallets, { ...novo, caixas: [] }] } : prev);
      setUnitizador(''); setEnderecoCodigo(''); setCodigo('');
      setMensagem(`Pallet ${novo.id} criado. Bipe o endereço.`);
    });
  }

  function biparEndereco() {
    if (!enderecoCodigo.trim() || !palletAberto || !atual || !pallet) return;
    void acao(async () => {
      const atualizado = await post<Pallet>(`/${atual.id}/pallets/${pallet.id}/endereco`, { codigo: enderecoCodigo.trim() });
      setAtual(prev => prev?.id === atual.id ? { ...prev, pallets: prev.pallets.map(p => p.id === pallet.id ? { ...p, ...atualizado, caixas: p.caixas } : p) } : prev);
      setEnderecoCodigo(''); setMensagem(`Endereço identificado: ${atualizado.endereco}. Bipe as caixas.`);
    });
  }

  function bipar() {
    if (!habilitado || !codigo.trim() || ocupado || !atual || !pallet) return;
    const leitura = codigo.trim();
    const destino = `/${atual.id}/pallets/${pallet.id}/caixas`;
    const inventarioId = atual.id;
    const leituraId = leituraRevisada.current ?? Array.from(crypto.getRandomValues(new Uint32Array(4)), n => n.toString(16).padStart(8, '0')).join('');
    leituraRevisada.current = null;
    const payload = { coletor_nome: coletorNome, leitura_id: leituraId, codigo: leitura, descricao: dados.descricao || undefined, validade: dados.validade || undefined, peso_kg: dados.peso_kg ? dados.peso_kg.replace(',', '.') : undefined };
    revisao.current++;
    setCodigo(''); setDados({ descricao: '', validade: '', peso_kg: '' });
    setPendentes(n => n + 1);
    input.current?.focus();
    fila.current = fila.current.then(async () => {
      try {
        const caixa = await post<Caixa>(destino, payload);
        revisao.current++;
        setAtual(prev => prev?.id === inventarioId ? { ...prev, pallets: prev.pallets.map(p => p.id === pallet.id ? { ...p, caixas: p.caixas.some(c => c.id === caixa.id) ? p.caixas : [caixa, ...p.caixas] } : p) } : prev);
        setMensagem(`Caixa salva no pallet ${pallet.id}: ${caixa.descricao} — ${kg(Number(caixa.peso_kg))} kg`);
        setFalhas(prev => prev.filter(f => f.leituraId !== leituraId));
      } catch (e) {
        setFalhas(prev => [...prev, { codigo: leitura, leituraId, inventarioId, palletId: pallet.id, mensagem: e instanceof Error ? e.message : 'Falha ao salvar. Confira a conexão e tente novamente.' }]);
      } finally { setPendentes(n => n - 1); }
    });
  }

  const todas = atual?.pallets.flatMap(p => p.caixas) ?? [];
  const pesoPallet = pallet?.caixas.reduce((s, c) => s + Number(c.peso_kg), 0) ?? 0;
  const orientacao = !atual ? 'Selecione ou crie um inventário para começar.' : atual.status !== 'aberto' ? 'Inventário encerrado. Os registros estão disponíveis para consulta.' : !pallet ? 'Bipe o unitizador para criar um pallet.' : pallet.status !== 'aberto' ? 'Pallet encerrado. Selecione um pallet aberto para continuar.' : !pallet.endereco_codigo ? 'Bipe o endereço para liberar a leitura das caixas.' : !pallet.unitizador ? 'Pallet antigo sem unitizador. Bipe um unitizador para criar um novo pallet.' : 'Toque no campo abaixo e bipe a etiqueta.';
  return <section className="inventory inv-workspace">
    <header className="inv-header">
      <div className="inv-title"><span className="inv-title-icon"><ClipboardList size={26} /></span><div><span className="inv-eyebrow">CONTAGEM DE ESTOQUE</span><h2>Inventário</h2></div></div>
      <div className="inv-sync"><span>{sincronizacao || 'Carregando inventários…'}</span><button aria-label="Atualizar inventário" title="Atualizar inventário" disabled={bloqueado} onClick={() => acao(() => carregar(atual?.id))}><RefreshCw size={18} /></button></div>
    </header>

    <details className="inv-device"><summary><Smartphone size={16} /><span>{coletorNome || 'Identificar coletor'}</span><span className="inv-device-edit">Editar identificação</span></summary><label>Nome deste coletor<input maxLength={120} value={coletorNome} disabled={bloqueado} onChange={e => setColetorNome(e.target.value)} /></label><p>Este nome acompanha as caixas registradas neste aparelho.</p></details>
    {erro && <div role="alert" className="alert">{erro}</div>}

    <div className="inv-selection-grid">
      <section className="inv-card inv-selection">
        <div className="inv-section-heading"><span className="inv-step">1</span><h3>Escolha o inventário</h3>{atual && <span className={`inv-badge ${atual.status}`}>{atual.status === 'aberto' ? 'Aberto' : 'Encerrado'}</span>}</div>
        <label><span className="inv-label">Inventário de trabalho</span><select disabled={bloqueado} value={atual?.id ?? ''} onChange={e => acao(() => carregar(Number(e.target.value) || undefined))}>
          <option value="" disabled>{lista.length ? 'Selecione um inventário' : 'Crie seu primeiro inventário'}</option>
          {lista.map(i => <option key={i.id} value={i.id}>Inventário {i.id} — {i.status}</option>)}
        </select></label>
        {atual && <p className="inv-hint">Operador: <strong>{atual.operador_nome || 'Não informado'}</strong></p>}
        <label>Nome do operador do novo inventário (obrigatório)<input value={operadorNome} maxLength={120} required disabled={bloqueado} onChange={e => setOperadorNome(e.target.value)} placeholder="Digite o nome de quem fará o inventário" /></label>
        <div className="inv-button-pair"><button className="inv-create" disabled={bloqueado || !operadorNome.trim()} onClick={() => acao(async () => { const novo = await post<Inventario>('', { operador_nome: operadorNome.trim() }); setOperadorNome(''); await carregar(novo.id); })}><Plus size={17} />Criar inventário</button><button className="inv-close" disabled={bloqueado || atual?.status !== 'aberto'} onClick={() => acao(async () => { await post(`/${atual!.id}/encerrar`); await carregar(atual!.id); })}><Square size={14} />Encerrar inventário</button></div>
      </section>
      <section className="inv-card inv-selection">
        <div className="inv-section-heading"><span className="inv-step">2</span><h3>Bipe o unitizador</h3>{pallet && <span className={`inv-badge ${pallet.status}`}>{pallet.status === 'aberto' ? 'Aberto' : 'Encerrado'}</span>}</div>
        <label><span className="inv-label">Destino das caixas</span><select disabled={bloqueado || !atual} value={palletId ?? ''} onChange={e => selecionarPallet(Number(e.target.value) || null)}>
          <option value="">Selecione ou crie seu pallet</option>
          {atual?.pallets.map(p => <option key={p.id} value={p.id}>Pallet {p.id} — {p.unitizador || 'Sem unitizador'} — {p.endereco || 'Sem endereço'} — {p.status}</option>)}
        </select></label>
        <form onSubmit={e => { e.preventDefault(); biparUnitizador(); }}>
          <label>Unitizador do novo pallet<input ref={unitizadorInput} value={unitizador} maxLength={120} disabled={bloqueado || atual?.status !== 'aberto'} autoComplete="off" onChange={e => setUnitizador(e.target.value)} onKeyDown={e => { if (e.key === 'Tab' && unitizador.trim()) { e.preventDefault(); biparUnitizador(); } }} placeholder="Bipe o unitizador e envie Enter" /></label>
          <button className="inv-create" disabled={bloqueado || atual?.status !== 'aberto' || !unitizador.trim()}>Salvar unitizador e criar pallet</button>
        </form>
        <p className="inv-hint">Cada bipe neste campo cria um novo pallet, mesmo se o número se repetir.</p>
        <div className="inv-button-pair"><button className="inv-create" disabled={bloqueado || atual?.status !== 'aberto'} onClick={() => unitizadorInput.current?.focus()}><Barcode size={17} />Bipar novo unitizador</button><button className="inv-close" disabled={bloqueado || !palletAberto} onClick={() => acao(async () => { await post(`/${atual!.id}/pallets/${palletId}/encerrar`); await carregar(atual!.id); })}><Square size={14} />Encerrar pallet</button></div>
      </section>
    </div>

    <section className="inv-card inv-selection">
      <div className="inv-section-heading"><span className="inv-step">3</span><h3>Bipe o endereço</h3></div>
      <p>Unitizador: <strong>{pallet?.unitizador || 'Não informado'}</strong> · Endereço: <strong>{pallet?.endereco || 'Aguardando leitura'}</strong></p>
      <form onSubmit={e => { e.preventDefault(); biparEndereco(); }}>
        <label>Código de barras do endereço<input ref={enderecoInput} value={enderecoCodigo} maxLength={120} autoComplete="off" disabled={bloqueado || !palletAberto || !pallet?.unitizador || !!pallet?.endereco_codigo} onChange={e => setEnderecoCodigo(e.target.value)} onKeyDown={e => { if (e.key === 'Tab' && enderecoCodigo.trim()) { e.preventDefault(); biparEndereco(); } }} placeholder="Bipe a etiqueta do endereço" /></label>
        <button className="inv-create" disabled={bloqueado || !palletAberto || !pallet?.unitizador || !!pallet?.endereco_codigo || !enderecoCodigo.trim()}>Identificar endereço</button>
      </form>
    </section>

    <form className={`inv-scan-panel ${habilitado ? 'is-ready' : ''}`} onSubmit={e => { e.preventDefault(); bipar(); }}>
      <div className="inv-scan-header"><div><span className="inv-eyebrow">DESTINO DA LEITURA</span><div className="inv-destination"><strong>{atual ? `Inventário ${atual.id}` : 'Nenhum inventário'}</strong><ChevronRight size={18} /><strong>{pallet ? `Pallet ${pallet.id}` : 'Nenhum pallet'}</strong></div></div><span className={`inv-badge ${habilitado ? 'aberto' : ''}`}>{habilitado ? 'Pronto para bipar' : 'Leitura desabilitada'}</span></div>
      <label htmlFor="inventory-scan" className="inv-scan-label"><Barcode size={24} />Leitura da caixa</label>
      <div className="inv-scan-controls"><input id="inventory-scan" ref={input} value={codigo} disabled={!habilitado || ocupado} onChange={e => { leituraRevisada.current = null; setCodigo(e.target.value); }} onKeyDown={e => { if (e.key === 'Tab' && codigo.trim()) { e.preventDefault(); bipar(); } }} autoComplete="off" autoCapitalize="off" spellCheck={false} placeholder="Toque aqui e bipe a etiqueta" /><button className="inv-add" disabled={!habilitado || ocupado || !codigo.trim()}><Plus size={20} />Adicionar caixa</button></div>
      <p className="inv-hint">{orientacao}</p>
      <div className={`inv-feedback ${pendentes ? 'pending' : mensagem ? 'success' : ''}`} role="status">{pendentes ? <><RefreshCw size={18} />Salvando {pendentes} leitura(s)…</> : mensagem ? <><CheckCircle2 size={20} /><span>{mensagem}</span></> : <><Barcode size={18} /><span>A confirmação de cada caixa aparecerá aqui.</span></>}</div>
      <details className="inv-manual"><summary>Etiqueta não identificada? Completar dados</summary><div className="inventory-selectors"><label>Nome do produto<input value={dados.descricao} onChange={e => setDados({ ...dados, descricao: e.target.value })} /></label><label>Validade<input type="date" value={dados.validade} onChange={e => setDados({ ...dados, validade: e.target.value })} /></label><label>Peso (kg)<input inputMode="decimal" value={dados.peso_kg} onChange={e => setDados({ ...dados, peso_kg: e.target.value })} /></label></div></details>
    </form>

    {falhas.length > 0 && <section className="inv-errors" role="alert"><h3>{falhas.length} leitura(s) não confirmada(s)</h3><p>Revise estas leituras antes de sair da página.</p>{falhas.map((f, i) => <article key={i}><div><strong>Inventário {f.inventarioId} · Pallet {f.palletId}</strong><code>{f.codigo}</code><p>{f.mensagem}</p></div><button disabled={bloqueado} onClick={() => acao(async () => { await carregar(f.inventarioId); selecionarPallet(f.palletId); setCodigo(f.codigo); leituraRevisada.current = f.leituraId; })}>Revisar leitura</button></article>)}</section>}

    <section className="inv-stats" aria-label="Resumo do pallet selecionado"><article><span className="inv-stat-icon"><Package size={22} /></span><div><span>Caixas neste pallet</span><strong>{pallet?.caixas.length ?? 0}</strong></div></article><article><span className="inv-stat-icon"><Weight size={22} /></span><div><span>Peso neste pallet</span><strong>{kg(pesoPallet)} <small>kg</small></strong></div></article></section>

    <section className="inv-card inv-box-section"><div className="inv-section-heading"><Boxes size={20} /><h3>Caixas do pallet {palletId ?? '—'}</h3><span className="inv-caption">Mais recentes primeiro</span></div>
      {!pallet?.caixas.length ? <div className="inv-empty"><Package size={32} /><strong>Nenhuma caixa registrada</strong><p>{habilitado ? 'Bipe a primeira etiqueta para iniciar a contagem.' : 'Escolha um inventário e um pallet aberto para começar.'}</p></div> : <div className="inv-box-list">{pallet.caixas.map(c => <article className="inv-box" key={c.id}><div className="inv-box-main"><span className="inv-box-number">#{c.id}</span><h4>{c.descricao}</h4><strong>{kg(Number(c.peso_kg))} <small>kg</small></strong></div><div className="inv-box-meta"><span>Validade <b>{c.validade}</b></span><span><Smartphone size={14} />{c.coletor_nome || 'Sem identificação'}</span></div><code>{c.codigo}</code></article>)}</div>}
    </section>

    <details className="inv-card inv-overview"><summary><ClipboardList size={20} /><strong>Resumo do inventário {atual?.id ?? '—'}</strong><span>{atual?.pallets.length ?? 0} pallets · {todas.length} caixas</span></summary><p className="inv-overview-total">Peso total: <strong>{kg(todas.reduce((s, c) => s + Number(c.peso_kg), 0))} kg</strong></p><div className="inv-pallet-grid">{atual?.pallets.map(p => <article className={p.id === palletId ? 'selected' : ''} key={p.id}><div><strong>Pallet {p.id}</strong><span className={`inv-badge ${p.status}`}>{p.status}</span></div><p>Unitizador: {p.unitizador || 'Não informado'} · {p.endereco || 'Sem endereço'}</p><p>{p.caixas.length} caixas · {kg(p.caixas.reduce((n, c) => n + Number(c.peso_kg), 0))} kg</p><span className="inv-caption">{[...new Set(p.caixas.map(c => c.coletor_nome).filter(Boolean))].join(', ') || 'Nenhuma identificação registrada'}</span>{p.id === palletId && <span className="inv-selected-label">Selecionado neste coletor</span>}</article>)}</div></details>
    <footer className="inv-footer">O coletor deve enviar Enter ou Tab após a etiqueta. Aguarde a confirmação de salvamento antes de sair.</footer>
  </section>;
}
