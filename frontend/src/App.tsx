import InventarioTab from './Inventario';
import ConsultaInventarios from './ConsultaInventarios';
import { useEffect, useMemo, useState } from 'react';
import type { ReactNode } from 'react';
import { Activity, Barcode, Boxes, Database, FileUp, Plus, Play, RefreshCcw, Search, Square, Trash2, Truck, Usb, Weight } from 'lucide-react';
import {
  Caixa,
  Coletor,
  EstoqueResumo,
  FiltrosPesquisa,
  LeituraColetor,
  Pesagem,
  Recebimento,
  ColetorOperacionalStatus,
  capturarLogcat,
  criarColetorAutomaticoUsb,
  finalizarColetorOperacional,
  finalizarConferencia,
  excluirColetor,
  health,
  importarLogcat,
  importarEstoqueAgregado,
  iniciarColetorOperacional,
  iniciarConferencia,
  listarCaixas,
  listarColetores,
  listarEstoque,
  listarLeituras,
  listarPesagens,
  listarRecebimentos,
  listarSaidas,
  reconectarIp,
  restaurarUsb,
  salvarColetor,
  statusColetorOperacional,
} from './api';

type Aba = 'consulta-inventarios' | 'inventario' | 'estoque' | 'saidas' | 'caixas' | 'pesagens' | 'recebimentos' | 'coletores';

const formatKg = (value: number | null | undefined) =>
  new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 3, maximumFractionDigits: 3 }).format(value ?? 0);

const formatDate = (value: string | null | undefined) => {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('pt-BR', {
    dateStyle: 'short',
    timeStyle: 'short',
  }).format(date);
};

const formatStatus = (value: string | null | undefined) => (value || '-').replaceAll('_', ' ');
const badgeClass = (value: string | null | undefined) => value || 'indefinido';

async function fileToBase64(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

export function App() {
  if (window.location.pathname.toLowerCase().includes('/consulta-inventarios')) {
    return <main className="content"><a className="inv-route-link" href="/inventario">Ir para contagem de inventário</a><ConsultaInventarios /></main>;
  }
  if (window.location.pathname.toLowerCase().includes('/inventario')) {
    return <main className="content"><a className="inv-route-link" href="/consulta-inventarios">Consultar inventários</a><InventarioTab /></main>;
  }
  const modoColetor = window.location.pathname.toLowerCase().includes('/coletor');
  return modoColetor ? <TelaColetor /> : <PainelDesktop />;
}

function TelaColetor() {
  const [status, setStatus] = useState<ColetorOperacionalStatus | null>(null);
  const [statusApi, setStatusApi] = useState<'online' | 'offline' | 'carregando'>('carregando');
  const [carregando, setCarregando] = useState(false);
  const [criandoColetorAuto, setCriandoColetorAuto] = useState(false);
  const [erro, setErro] = useState('');
  const [mensagem, setMensagem] = useState('');

  async function carregarStatus() {
    try {
      await health();
      const dados = await statusColetorOperacional();
      setStatus(dados);
      setStatusApi('online');
      setErro(dados.ok ? '' : dados.erro || 'Coletor nao encontrado.');
    } catch (error) {
      setStatusApi('offline');
      setErro(error instanceof Error ? error.message : 'Falha ao carregar status.');
    }
  }

  useEffect(() => {
    carregarStatus();
    const timer = window.setInterval(carregarStatus, 5000);
    return () => window.clearInterval(timer);
  }, []);

  async function executarConferencia(acao: typeof iniciarColetorOperacional) {
    setCarregando(true);
    setErro('');
    setMensagem('');
    try {
      const resposta = await acao();
      setStatus(resposta);
      if (!resposta.ok) {
        setErro(resposta.erro || 'Acao nao concluida.');
        return;
      }
      if (resposta.captura) {
        setMensagem(`Conferencia finalizada. Baixadas: ${resposta.captura.baixadas ?? 0}. Nao encontradas: ${resposta.captura.nao_encontradas ?? 0}.`);
      } else {
        setMensagem(resposta.aviso || 'Conferencia iniciada. Pode bipar as caixas no WMS.');
      }
    } catch (error) {
      setErro(error instanceof Error ? error.message : 'Falha ao executar acao.');
    } finally {
      setCarregando(false);
    }
  }

  const ativo = !!status?.sessao;
  const coletorNome = status?.coletor?.nome ?? 'Coletor';

  return (
    <main className="collector-screen">
      <section className="collector-status">
        <div>
          <span className={`status-dot ${statusApi}`} />
          <span className="status-text">{statusApi === 'online' ? 'Conectado' : 'Sem conexao'}</span>
        </div>
        <strong className={ativo ? 'active' : ''}>{ativo ? 'Conferencia ativa' : 'Sem conferencia ativa'}</strong>
        <span>{ativo ? `Sessao #${status?.sessao?.id}` : coletorNome}</span>
        <small>{ativo ? `Inicio: ${formatDate(status?.sessao?.inicio)}` : 'Selecao de estoque WMS'}</small>
      </section>

      {erro && <div className="alert collector-alert">{erro}</div>}
      {mensagem && <div className="notice collector-alert">{mensagem}</div>}

      <section className="collector-actions">
        <button
          className="collector-action start"
          onClick={() => executarConferencia(iniciarColetorOperacional)}
          disabled={carregando || ativo || statusApi !== 'online' || status?.ok === false}
        >
          <Play size={46} />
          <span>Iniciar conferencia</span>
        </button>
        <button
          className="collector-action finish"
          onClick={() => executarConferencia(finalizarColetorOperacional)}
          disabled={carregando || !ativo || statusApi !== 'online'}
        >
          <Square size={46} />
          <span>Finalizar</span>
        </button>
      </section>
    </main>
  );
}

function PainelDesktop() {
  const [aba, setAba] = useState<Aba>(window.location.hash === '#consulta-inventarios' ? 'consulta-inventarios' : window.location.hash === '#inventario' ? 'inventario' : 'estoque');
  const [filtros, setFiltros] = useState<FiltrosPesquisa>({ busca: '', nome: '', validade: '', peso: '' });
  const [estoque, setEstoque] = useState<EstoqueResumo[]>([]);
  const [caixas, setCaixas] = useState<Caixa[]>([]);
  const [pesagens, setPesagens] = useState<Pesagem[]>([]);
  const [recebimentos, setRecebimentos] = useState<Recebimento[]>([]);
  const [coletores, setColetores] = useState<Coletor[]>([]);
  const [saidas, setSaidas] = useState<LeituraColetor[]>([]);
  const [leituras, setLeituras] = useState<LeituraColetor[]>([]);
  const [sessaoSelecionada, setSessaoSelecionada] = useState<number | null>(null);
  const [totais, setTotais] = useState({ estoque: 0, saidas: 0, caixas: 0, pesagens: 0, recebimentos: 0, coletores: 0 });
  const [statusApi, setStatusApi] = useState<'online' | 'offline' | 'carregando'>('carregando');
  const [carregando, setCarregando] = useState(false);
  const [criandoColetorAuto, setCriandoColetorAuto] = useState(false);
  const [erro, setErro] = useState('');
  const [mensagem, setMensagem] = useState('');
  const [novoColetor, setNovoColetor] = useState({ nome: '', serial: '', ip_atual: '' });
  const [operador, setOperador] = useState('');

  const metricas = useMemo(() => {
    const pesoTotal = estoque.reduce((acc, item) => acc + item.peso_total_kg, 0);
    const caixas = estoque.reduce((acc, item) => acc + item.qtd_etiquetas, 0);
    return { pesoTotal, caixas, produtos: totais.estoque };
  }, [estoque, totais.estoque]);

  async function carregar() {
    setCarregando(true);
    setErro('');
    try {
      await health();
      setStatusApi('online');
      const [dadosEstoque, dadosSaidas, dadosCaixas, dadosPesagens, dadosRecebimentos, dadosColetores] = await Promise.all([
        listarEstoque(filtros, 0),
        listarSaidas(filtros, 0),
        listarCaixas(filtros, 0),
        listarPesagens(filtros, 0),
        listarRecebimentos(),
        listarColetores(),
      ]);
      setEstoque(dadosEstoque.items);
      setSaidas(dadosSaidas.items);
      setCaixas(dadosCaixas.items);
      setPesagens(dadosPesagens.items);
      setRecebimentos(dadosRecebimentos.items);
      setColetores(dadosColetores.items);
      const sessaoAtiva = dadosColetores.items.find((coletor) => coletor.sessao_ativa_id)?.sessao_ativa_id;
      if (sessaoAtiva && aba !== 'inventario' && aba !== 'consulta-inventarios') {
        await capturarLogcat(sessaoAtiva);
        const coletoresAtualizados = await listarColetores();
        setColetores(coletoresAtualizados.items);
      }
      const sessaoParaMostrar =
        sessaoAtiva ??
        dadosColetores.items.find((coletor) => coletor.ultima_sessao_id)?.ultima_sessao_id ??
        sessaoSelecionada;
      if (sessaoParaMostrar) {
        const dadosLeituras = await listarLeituras(sessaoParaMostrar);
        setLeituras(dadosLeituras.items);
        setSessaoSelecionada(sessaoParaMostrar);
      }
      setTotais({
        estoque: dadosEstoque.total,
        saidas: dadosSaidas.total,
        caixas: dadosCaixas.total,
        pesagens: dadosPesagens.total,
        recebimentos: dadosRecebimentos.total,
        coletores: dadosColetores.total,
      });
    } catch (error) {
      setStatusApi('offline');
      setErro(error instanceof Error ? error.message : 'Falha ao carregar dados.');
    } finally {
      setCarregando(false);
    }
  }

  useEffect(() => {
    if (aba === 'inventario' || aba === 'consulta-inventarios') return;
    carregar();
    const timer = window.setInterval(carregar, 10000);
    return () => window.clearInterval(timer);
  }, [aba]);

  async function executarAcao(acao: () => Promise<unknown>, textoSucesso: string) {
    setErro('');
    setMensagem('');
    try {
      const resposta = await acao();
      if (resposta && typeof resposta === 'object' && 'ok' in resposta && !resposta.ok) {
        setErro('erro' in resposta && typeof resposta.erro === 'string' ? resposta.erro : 'Ação não concluída.');
        return;
      }
      setMensagem(
        resposta && typeof resposta === 'object' && 'aviso' in resposta && typeof resposta.aviso === 'string'
          ? resposta.aviso
          : textoSucesso,
      );
      await carregar();
    } catch (error) {
      setErro(error instanceof Error ? error.message : 'Falha ao executar ação.');
    }
  }

  async function selecionarSessao(sessaoId: number | null) {
    setSessaoSelecionada(sessaoId);
    if (!sessaoId) {
      setLeituras([]);
      return;
    }
    const dados = await listarLeituras(sessaoId);
    setLeituras(dados.items);
  }

  async function cadastrarColetor() {
    if (!novoColetor.nome || !novoColetor.serial) {
      setErro('Informe nome e serial do coletor.');
      return;
    }
    await executarAcao(
      () => salvarColetor({ ...novoColetor, ip_atual: novoColetor.ip_atual || undefined }),
      'Coletor salvo.',
    );
    setNovoColetor({ nome: '', serial: '', ip_atual: '' });
  }

  async function cadastrarColetorAutomatico() {
    setCriandoColetorAuto(true);
    setErro('');
    setMensagem('Procurando coletor conectado no USB...');
    try {
      const resposta = await criarColetorAutomaticoUsb(novoColetor.serial || undefined);
      if (!resposta.ok) {
        const detalhe = resposta.devices ? ` Dispositivos ADB: ${resposta.devices.trim() || 'nenhum'}` : '';
        setErro(`${resposta.erro || 'Nao foi possivel criar o coletor automaticamente.'}${detalhe}`);
        setMensagem('');
        return;
      }

      const coletor = resposta.coletor;
      const acao = resposta.criado ? 'cadastrado' : 'atualizado';
      setMensagem(
        `Coletor ${acao}: ${coletor?.nome || resposta.nome || '-'} | Serial ${resposta.serial || coletor?.serial || '-'} | IP ${resposta.ip_atual || coletor?.ip_atual || '-'}.`,
      );
      setNovoColetor({ nome: '', serial: '', ip_atual: '' });
      const coletoresAtualizados = await listarColetores();
      setColetores(coletoresAtualizados.items);
      setTotais((atual) => ({ ...atual, coletores: coletoresAtualizados.total }));
    } catch (error) {
      setErro(error instanceof Error ? error.message : 'Falha ao criar coletor automaticamente.');
      setMensagem('');
    } finally {
      setCriandoColetorAuto(false);
    }
  }

  async function importarArquivoLog(coletor: Coletor, file: File | null) {
    if (!file) return;
    if (!coletor.sessao_ativa_id) {
      setErro('Inicie uma conferência antes de importar o logcat.');
      return;
    }
    const texto = await file.text();
    await executarAcao(
      () => importarLogcat(coletor.sessao_ativa_id as number, texto),
      'Logcat importado e processado.',
    );
  }

  async function importarArquivoEstoque(file: File | null) {
    if (!file) return;
    setErro('');
    setMensagem('');
    try {
      const arquivoBase64 = await fileToBase64(file);
      const resposta = await importarEstoqueAgregado({ nome_arquivo: file.name, arquivo_base64: arquivoBase64 });
      if (!resposta.ok) {
        setErro(resposta.erro || 'Falha ao importar estoque.');
        return;
      }
      setMensagem(`Estoque importado: ${resposta.linhas_importadas ?? 0} linhas.`);
      await carregar();
    } catch (error) {
      setErro(error instanceof Error ? error.message : 'Falha ao importar estoque.');
    }
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand-mark">PC</div>
        <div>
          <h1>PesoCtrl</h1>
          <p>Controle de estoque por etiqueta</p>
        </div>
        <nav>
          <button className={aba === 'consulta-inventarios' ? 'active' : ''} onClick={() => setAba('consulta-inventarios')}>
            <Search size={18} /> Consulta de inventários
          </button>
          <button className={aba === 'inventario' ? 'active' : ''} onClick={() => setAba('inventario')}>
            <Boxes size={18} /> Invent?rio
          </button>
          <button className={aba === 'estoque' ? 'active' : ''} onClick={() => setAba('estoque')}>
            <Boxes size={18} /> Estoque
          </button>
          <button className={aba === 'saidas' ? 'active' : ''} onClick={() => setAba('saidas')}>
            <Truck size={18} /> Saídas
          </button>
          <button className={aba === 'caixas' ? 'active' : ''} onClick={() => setAba('caixas')}>
            <Barcode size={18} /> Caixas
          </button>
          <button className={aba === 'pesagens' ? 'active' : ''} onClick={() => setAba('pesagens')}>
            <Weight size={18} /> Pesagens
          </button>
          <button className={aba === 'recebimentos' ? 'active' : ''} onClick={() => setAba('recebimentos')}>
            <Truck size={18} /> Recebimentos
          </button>
          <button className={aba === 'coletores' ? 'active' : ''} onClick={() => setAba('coletores')}>
            <Usb size={18} /> Coletores
          </button>
        </nav>
      </aside>

      <section className="content">
        {aba !== 'inventario' && aba !== 'consulta-inventarios' && <>
        <header className="topbar">
          <div>
            <span className={`status-dot ${statusApi}`} />
            <span className="status-text">{statusApi === 'online' ? 'API online' : 'API offline'}</span>
          </div>
          <div className="actions">
            <label className="search">
              <Search size={18} />
              <input
                value={filtros.busca}
                onChange={(event) => setFiltros((atual) => ({ ...atual, busca: event.target.value }))}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') carregar();
                }}
                placeholder="Buscar produto, caixa ou impressora"
              />
            </label>
            <button className="icon-button" onClick={carregar} disabled={carregando} title="Atualizar">
              <RefreshCcw size={18} />
            </button>
          </div>
        </header>

        <section className="filters">
          <label>
            <span>Nome</span>
            <input
              value={filtros.nome}
              onChange={(event) => setFiltros((atual) => ({ ...atual, nome: event.target.value }))}
              onKeyDown={(event) => {
                if (event.key === 'Enter') carregar();
              }}
              placeholder="Produto"
            />
          </label>
          <label>
            <span>Validade</span>
            <input
              value={filtros.validade}
              onChange={(event) => setFiltros((atual) => ({ ...atual, validade: event.target.value }))}
              onKeyDown={(event) => {
                if (event.key === 'Enter') carregar();
              }}
              placeholder="dd/mm/aaaa"
            />
          </label>
          <label>
            <span>Peso</span>
            <input
              value={filtros.peso}
              onChange={(event) => setFiltros((atual) => ({ ...atual, peso: event.target.value.replace(',', '.') }))}
              onKeyDown={(event) => {
                if (event.key === 'Enter') carregar();
              }}
              placeholder="25.000"
              inputMode="decimal"
            />
          </label>
          <button className="filter-button" onClick={carregar} disabled={carregando}>
            <Search size={17} /> Pesquisar
          </button>
        </section>

        <section className="metrics">
          <Metric icon={<Database size={21} />} label="Produtos/validade" value={metricas.produtos.toString()} />
          <Metric icon={<Boxes size={21} />} label="Caixas" value={totais.caixas.toString()} />
          <Metric icon={<Weight size={21} />} label="Peso total" value={`${formatKg(metricas.pesoTotal)} KG`} />
          <Metric icon={<Activity size={21} />} label="Pesagens" value={totais.pesagens.toString()} />
        </section>

        {erro && <div className="alert">{erro}</div>}
        {mensagem && <div className="notice">{mensagem}</div>}

        </>}
        <div hidden={aba !== 'inventario'}><InventarioTab /></div>
        {aba === 'consulta-inventarios' && <ConsultaInventarios />}

        {aba === 'estoque' && (
          <Panel title="Estoque por produto e validade" count={totais.estoque}>
            <div className="panel-toolbar">
              <label className="filter-button secondary import-button">
                <FileUp size={17} /> Importar estoque WMS
                <input
                  type="file"
                  accept=".xlsx"
                  onChange={(event) => importarArquivoEstoque(event.target.files?.[0] ?? null)}
                />
              </label>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Produto</th>
                  <th>Validade</th>
                  <th>Caixas</th>
                  <th>Peso total</th>
                  <th>Leituras</th>
                </tr>
              </thead>
              <tbody>
                {estoque.map((item) => (
                  <tr key={`${item.descricao}-${item.validade}`}>
                    <td>{item.descricao || '-'}</td>
                    <td>{item.validade || '-'}</td>
                    <td>{item.qtd_etiquetas}</td>
                    <td>{formatKg(item.peso_total_kg)} KG</td>
                    <td>{item.leituras}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Panel>
        )}

        {aba === 'saidas' && (
          <Panel title="Caixas bipadas na seleção" count={totais.saidas}>
            <table>
              <thead>
                <tr>
                  <th>Data/hora</th>
                  <th>Coletor</th>
                  <th>Produto</th>
                  <th>Peso lido</th>
                  <th>Validade</th>
                  <th>Status</th>
                  <th>Origem da baixa</th>
                  <th>Código lido</th>
                  <th>Caixa</th>
                  <th>Agregado</th>
                </tr>
              </thead>
              <tbody>
                {saidas.map((item) => (
                  <tr key={item.id}>
                    <td>{formatDate(item.data_hora)}</td>
                    <td>{item.coletor || '-'}</td>
                    <td>{item.produto_codigo ? `${item.produto_codigo} - ${item.produto_descricao || ''}` : item.produto_descricao || '-'}</td>
                    <td>{item.peso_lido_kg !== null ? `${formatKg(item.peso_lido_kg)} KG` : '-'}</td>
                    <td>{item.validade_lida || '-'}</td>
                    <td><span className={`badge ${badgeClass(item.status)}`}>{formatStatus(item.status)}</span></td>
                    <td>{item.regra_baixa ? item.regra_baixa.replaceAll('_', ' ') : '-'}</td>
                    <td><code>{item.codigo_lido}</code></td>
                    <td>{item.caixa_id ?? '-'}</td>
                    <td>{item.estoque_agregado_id ?? '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Panel>
        )}

        {aba === 'caixas' && (
          <Panel title="Caixas registradas" count={totais.caixas}>
            <table>
              <thead>
                <tr>
                  <th>ID da caixa</th>
                  <th>Código original</th>
                  <th>Data/hora</th>
                  <th>Produto</th>
                  <th>Pallet</th>
                  <th>Impressora</th>
                  <th>Peso</th>
                  <th>Validade</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {caixas.map((item) => (
                  <tr key={item.id}>
                    <td><code>{item.codigo_caixa}</code></td>
                    <td><code>{item.codigo_barras_original || '-'}</code></td>
                    <td>{formatDate(item.data_hora)}</td>
                    <td>{item.descricao || '-'}</td>
                    <td>{item.numero_pallet ?? '-'}</td>
                    <td>{item.impressora}</td>
                    <td>{formatKg(item.peso_kg)} KG</td>
                    <td>{item.validade || '-'}</td>
                    <td><span className={`badge ${badgeClass(item.status)}`}>{formatStatus(item.status)}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Panel>
        )}

        {aba === 'pesagens' && (
          <Panel title="Histórico de pesagens" count={totais.pesagens}>
            <table>
              <thead>
                <tr>
                  <th>Data/hora</th>
                  <th>Produto</th>
                  <th>Pallet</th>
                  <th>Impressora</th>
                  <th>Peso</th>
                  <th>Qtd.</th>
                  <th>Total</th>
                  <th>Validade</th>
                </tr>
              </thead>
              <tbody>
                {pesagens.map((item) => (
                  <tr key={item.id}>
                    <td>{formatDate(item.data_hora)}</td>
                    <td>{item.descricao || '-'}</td>
                    <td>{item.numero_pallet ?? '-'}</td>
                    <td>{item.impressora}</td>
                    <td>{formatKg(item.peso_kg)} KG</td>
                    <td>{item.qtd_etiquetas}</td>
                    <td>{formatKg(item.peso_total_kg)} KG</td>
                    <td>{item.validade || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Panel>
        )}

        {aba === 'recebimentos' && (
          <Panel title="Recebimentos" count={totais.recebimentos}>
            <table>
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Início</th>
                  <th>Fim</th>
                  <th>Status</th>
                  <th>Pallets</th>
                </tr>
              </thead>
              <tbody>
                {recebimentos.map((item) => (
                  <tr key={item.id}>
                    <td>{item.recebimento_id}</td>
                    <td>{formatDate(item.inicio)}</td>
                    <td>{formatDate(item.fim)}</td>
                    <td><span className={`badge ${badgeClass(item.status)}`}>{formatStatus(item.status)}</span></td>
                    <td>{item.total_pallets}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Panel>
        )}

        {aba === 'coletores' && (
          <Panel title="Coletores Honeywell" count={totais.coletores}>
            <div className="collector-form">
              <label>
                <span>Nome</span>
                <input
                  value={novoColetor.nome}
                  onChange={(event) => setNovoColetor((atual) => ({ ...atual, nome: event.target.value }))}
                  placeholder="EXP-01"
                />
              </label>
              <label>
                <span>Serial</span>
                <input
                  value={novoColetor.serial}
                  onChange={(event) => setNovoColetor((atual) => ({ ...atual, serial: event.target.value }))}
                  placeholder="Serial do EDA61K"
                />
              </label>
              <label>
                <span>IP atual</span>
                <input
                  value={novoColetor.ip_atual}
                  onChange={(event) => setNovoColetor((atual) => ({ ...atual, ip_atual: event.target.value }))}
                  placeholder="192.168.1.238"
                />
              </label>
              <label>
                <span>Operador</span>
                <input value={operador} onChange={(event) => setOperador(event.target.value)} placeholder="Nome" />
              </label>
              <button className="filter-button" onClick={cadastrarColetor}>Salvar coletor</button>
              <button
                className="filter-button"
                onClick={cadastrarColetorAutomatico}
                disabled={criandoColetorAuto}
                title="Detectar coletor conectado via USB e cadastrar automaticamente"
              >
                <Plus size={17} /> {criandoColetorAuto ? 'Procurando...' : 'Criar auto'}
              </button>
              <button
                className="filter-button secondary"
                onClick={() => executarAcao(() => restaurarUsb(), 'Restauração USB executada.')}
                disabled={carregando}
              >
                <Usb size={17} /> Restaurar USB
              </button>
            </div>

            <table>
              <thead>
                <tr>
                  <th>Coletor</th>
                  <th>Serial</th>
                  <th>IP</th>
                  <th>Status</th>
                  <th>Último bipe</th>
                  <th>Código</th>
                  <th>Ações</th>
                </tr>
              </thead>
              <tbody>
                {coletores.map((item) => (
                  <tr key={item.id}>
                    <td>{item.nome}</td>
                    <td><code>{item.serial}</code></td>
                    <td>{item.ip_atual || '-'}</td>
                    <td><span className={`badge ${badgeClass(item.status)}`}>{formatStatus(item.status)}</span></td>
                    <td>{formatDate(item.ultimo_bipe)}</td>
                    <td><code>{item.ultimo_codigo_lido || '-'}</code></td>
                    <td>
                      <div className="row-actions">
                        <button
                          className="mini-button"
                          onClick={() => executarAcao(
                            () => iniciarConferencia(item.id, { operador: operador || undefined }),
                            'Conferência iniciada e logcat limpo.',
                          )}
                          disabled={!!item.sessao_ativa_id}
                          title="Iniciar conferência"
                        >
                          <Play size={15} />
                        </button>
                        <label className={`mini-button file-action ${!item.sessao_ativa_id ? 'disabled' : ''}`} title="Importar TXT do logcat">
                          <FileUp size={15} />
                          <input
                            type="file"
                            accept=".txt,.log"
                            disabled={!item.sessao_ativa_id}
                            onChange={(event) => importarArquivoLog(item, event.target.files?.[0] ?? null)}
                          />
                        </label>
                        <button
                          className="mini-button"
                          onClick={() => item.sessao_ativa_id && executarAcao(
                            () => capturarLogcat(item.sessao_ativa_id as number),
                            'Logcat capturado do coletor.',
                          )}
                          disabled={!item.sessao_ativa_id}
                          title="Capturar logcat agora"
                        >
                          <RefreshCcw size={15} />
                        </button>
                        <button
                          className="mini-button"
                          onClick={() => item.sessao_ativa_id && executarAcao(
                            () => finalizarConferencia(item.sessao_ativa_id as number),
                            'Conferência finalizada.',
                          )}
                          disabled={!item.sessao_ativa_id}
                          title="Finalizar conferência"
                        >
                          <Square size={15} />
                        </button>
                        <button
                          className="mini-button"
                          onClick={() => executarAcao(
                            () => reconectarIp(item.id),
                            'Coletor reconectado por IP.',
                          )}
                          title="Reconectar por IP"
                        >
                          <Usb size={15} />
                        </button>
                        <button
                          className="mini-button"
                          onClick={() => executarAcao(() => excluirColetor(item.id), 'Coletor excluído.')}
                          disabled={!!item.sessao_ativa_id}
                          title="Excluir coletor"
                        >
                          <Trash2 size={15} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="readings-panel">
              <div className="panel-head compact">
                <h2>Leituras importadas</h2>
                <span>{sessaoSelecionada ? `Sessão #${sessaoSelecionada}` : 'Nenhuma sessão selecionada'}</span>
              </div>
              <table>
                <thead>
                  <tr>
                    <th>Data/hora</th>
                    <th>Código lido</th>
                    <th>Status</th>
                    <th>Caixa</th>
                  </tr>
                </thead>
                <tbody>
                  {leituras.map((item) => (
                    <tr key={item.id}>
                      <td>{formatDate(item.data_hora)}</td>
                      <td><code>{item.codigo_lido}</code></td>
                      <td><span className={`badge ${badgeClass(item.status)}`}>{formatStatus(item.status)}</span></td>
                      <td>{item.caixa_id ?? '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>
        )}
      </section>
    </main>
  );
}

function Metric({ icon, label, value }: { icon: ReactNode; label: string; value: string }) {
  return (
    <article className="metric">
      <div className="metric-icon">{icon}</div>
      <div>
        <span>{label}</span>
        <strong>{value}</strong>
      </div>
    </article>
  );
}

function Panel({ title, count, children }: { title: string; count: number; children: ReactNode }) {
  return (
    <section className="panel">
      <div className="panel-head">
        <h2>{title}</h2>
        <span>{count} registros</span>
      </div>
      <div className="table-wrap">{children}</div>
    </section>
  );
}
