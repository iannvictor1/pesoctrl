const API_BASE =
  import.meta.env.VITE_API_BASE ??
  `${window.location.protocol}//${window.location.hostname}:8001`;

type QueryValue = string | number | undefined | null;

export async function exportarInventario(id: number): Promise<void> {
  const response = await fetch(`${API_BASE}/api/inventarios/${id}/exportar`);
  if (!response.ok) {
    const body = await response.json().catch(() => null);
    throw new Error(typeof body?.detail === 'string' ? body.detail : 'Não foi possível exportar a planilha. Tente novamente.');
  }
  const url = URL.createObjectURL(await response.blob());
  const link = document.createElement('a');
  link.href = url;
  link.download = `inventario_${id}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}

export async function request<T>(
  path: string,
  params: Record<string, QueryValue> = {},
  init: RequestInit = {},
): Promise<T> {
  const url = new URL(`${API_BASE}${path}`);
  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(key, String(value));
    }
  });

  const response = await fetch(url, init);
  if (!response.ok) {
    const body = await response.json().catch(() => null);
    if (typeof body?.detail === 'string') throw new Error(body.detail);
    throw new Error(`Erro ${response.status}: ${response.statusText}`);
  }
  return response.json();
}

export type Lista<T> = {
  total: number;
  items: T[];
};

export type Recebimento = {
  id: number;
  recebimento_id: string;
  inicio: string;
  fim: string | null;
  status: string;
  total_pallets: number;
};

export type EstoqueResumo = {
  descricao: string;
  validade: string | null;
  qtd_etiquetas: number;
  peso_total_kg: number;
  leituras: number;
};

export type Pesagem = {
  id: number;
  data_hora: string;
  recebimento_id: string | null;
  numero_pallet: number | null;
  impressora: string;
  descricao: string;
  peso_kg: number | null;
  qtd_etiquetas: number;
  peso_total_kg: number | null;
  validade: string | null;
};

export type Caixa = {
  id: number;
  codigo_caixa: string;
  codigo_barras_original: string | null;
  data_hora: string;
  recebimento_id: string | null;
  numero_pallet: number | null;
  impressora: string;
  descricao: string;
  peso_kg: number | null;
  validade: string | null;
  status: string;
};

export type Coletor = {
  id: number;
  nome: string;
  serial: string;
  ip_atual: string | null;
  status: string;
  ultimo_contato: string | null;
  ultimo_bipe: string | null;
  ultimo_codigo_lido: string | null;
  sessao_ativa_id: number | null;
  ultima_sessao_id: number | null;
};

export type SessaoConferencia = {
  id: number;
  coletor_id: number;
  inicio: string;
  fim: string | null;
  status: string;
  operador: string | null;
  observacao: string | null;
  log_limpo: boolean;
  total_leituras: number;
  baixadas: number;
  duplicadas: number;
  nao_encontradas: number;
  ignoradas: number;
};

export type LeituraColetor = {
  id: number;
  data_hora: string;
  codigo_lido: string;
  status: string;
  caixa_id: number | null;
  estoque_agregado_id: number | null;
  regra_baixa: string | null;
  produto_codigo: string | null;
  produto_descricao: string | null;
  peso_lido_kg: number | null;
  validade_lida: string | null;
  peso_baixado_kg: number | null;
  sessao_id?: number | null;
  coletor?: string | null;
};

export type RecebimentoGlobalStatus = {
  recebimento_ativo: boolean;
  recebimento_id: string | null;
  inicio_recebimento: string | null;
  contador_pallets: number;
};

export type ColetorOperacionalStatus = {
  ok: boolean;
  erro?: string;
  aviso?: string;
  coletor?: Coletor;
  sessao?: SessaoConferencia | null;
};

export type ColetorOperacionalResultado = ColetorOperacionalStatus & {
  captura?: {
    total?: number;
    linhas_extraidas?: number;
    baixadas?: number;
    duplicadas?: number;
    nao_encontradas?: number;
    ignoradas?: number;
  };
};

export type FiltrosPesquisa = {
  busca?: string;
  nome?: string;
  validade?: string;
  peso?: string;
};

export function health() {
  return request<{ status: string; time: string }>('/api/health');
}

export function listarRecebimentos() {
  return request<Lista<Recebimento>>('/api/recebimentos', { limit: 20 });
}

export function listarEstoque(filtros: FiltrosPesquisa, offset: number) {
  return request<Lista<EstoqueResumo>>('/api/estoque/resumo', { ...filtros, offset, limit: 100 });
}

export function listarPesagens(filtros: FiltrosPesquisa, offset: number) {
  return request<Lista<Pesagem>>('/api/pesagens', { ...filtros, offset, limit: 100 });
}

export function listarCaixas(filtros: FiltrosPesquisa, offset: number) {
  return request<Lista<Caixa>>('/api/caixas', { ...filtros, offset, limit: 100 });
}

function post<T>(path: string, body: unknown): Promise<T> {
  return request<T>(path, {}, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
}

function del<T>(path: string): Promise<T> {
  return request<T>(path, {}, { method: 'DELETE' });
}

export function listarColetores() {
  return request<Lista<Coletor>>('/api/coletores');
}

export function salvarColetor(payload: { nome: string; serial: string; ip_atual?: string }) {
  return post<Coletor>('/api/coletores', payload);
}

export function excluirColetor(coletorId: number) {
  return del<{ ok: boolean; erro?: string }>(`/api/coletores/${coletorId}`);
}

export function iniciarConferencia(coletorId: number, payload: { operador?: string; observacao?: string }) {
  return post<{ ok: boolean; erro?: string; aviso?: string; sessao?: SessaoConferencia }>(
    `/api/coletores/${coletorId}/iniciar-conferencia`,
    payload,
  );
}

export function finalizarConferencia(sessaoId: number) {
  return post<{ ok: boolean; erro?: string; sessao?: SessaoConferencia }>(
    `/api/sessoes-conferencia/${sessaoId}/finalizar`,
    {},
  );
}

export function importarLogcat(sessaoId: number, conteudo: string) {
  return post<{
    ok: boolean;
    erro?: string;
    total: number;
    linhas_extraidas: number;
    baixadas: number;
    duplicadas: number;
    nao_encontradas: number;
    ignoradas: number;
    sessao?: SessaoConferencia;
  }>(`/api/sessoes-conferencia/${sessaoId}/importar-logcat`, { conteudo });
}

export function capturarLogcat(sessaoId: number) {
  return post<{
    ok: boolean;
    erro?: string;
    total: number;
    linhas_extraidas: number;
    baixadas: number;
    duplicadas: number;
    nao_encontradas: number;
    ignoradas: number;
    sessao?: SessaoConferencia;
  }>(`/api/sessoes-conferencia/${sessaoId}/capturar-logcat`, {});
}

export function listarLeituras(sessaoId: number) {
  return request<Lista<LeituraColetor>>(`/api/sessoes-conferencia/${sessaoId}/leituras`, { limit: 50 });
}

export function listarSaidas(filtros: FiltrosPesquisa, offset: number) {
  return request<Lista<LeituraColetor>>('/api/saidas', { ...filtros, offset, limit: 100 });
}

export function limparLogColetor(coletorId: number) {
  return post<{ ok: boolean; erro?: string; stdout?: string; stderr?: string }>(`/api/coletores/${coletorId}/limpar-log`, {});
}

export function reconectarIp(coletorId: number) {
  return post<{ ok: boolean; erro?: string; stdout?: string; stderr?: string }>(`/api/coletores/${coletorId}/reconectar-ip`, {});
}

export function restaurarUsb(serial?: string) {
  return post<{ ok: boolean; erro?: string; serial?: string; ip_atual?: string; stdout?: string; stderr?: string }>(
    '/api/coletores/restaurar-usb',
    { serial },
  );
}

export function criarColetorAutomaticoUsb(serial?: string) {
  return post<{
    ok: boolean;
    erro?: string;
    criado?: boolean;
    serial?: string;
    ip_atual?: string;
    nome?: string;
    coletor?: Coletor;
    devices?: string;
    stdout?: string;
    stderr?: string;
  }>('/api/coletores/criar-automatico-usb', { serial });
}

export function statusRecebimentoGlobal() {
  return request<RecebimentoGlobalStatus>('/api/recebimento-global/status');
}

export function iniciarRecebimentoGlobal() {
  return post<{ ok: boolean; mensagem: string; status: RecebimentoGlobalStatus }>(
    '/api/recebimento-global/iniciar',
    {},
  );
}

export function finalizarRecebimentoGlobal() {
  return post<{ ok: boolean; mensagem: string; status: RecebimentoGlobalStatus }>(
    '/api/recebimento-global/finalizar',
    {},
  );
}

export function statusColetorOperacional() {
  return request<ColetorOperacionalStatus>('/api/coletor-operacional/status');
}

export function iniciarColetorOperacional() {
  return post<ColetorOperacionalResultado>('/api/coletor-operacional/iniciar', {});
}

export function finalizarColetorOperacional() {
  return post<ColetorOperacionalResultado>('/api/coletor-operacional/finalizar', {});
}

export function importarEstoqueAgregado(payload: { nome_arquivo?: string; arquivo_base64: string }) {
  return post<{
    ok: boolean;
    erro?: string;
    importacao_id?: string;
    linhas_importadas?: number;
    linhas_ignoradas?: number;
  }>('/api/estoque-agregado/importar', payload);
}
