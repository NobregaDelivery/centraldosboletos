/**
 * ORGANIZADOR DE BOLETOS — Google Apps Script
 * --------------------------------------------
 * Lê o seu Gmail pela API oficial do Google, encontra os boletos dos
 * fornecedores cadastrados na aba "Fornecedores", extrai valor, vencimento
 * e linha digitável (do corpo do e-mail ou do PDF anexo) e registra tudo
 * na aba "Contas". Todo dia envia um e-mail de alerta com o que está
 * vencendo e, se quiser, cria um lembrete no Google Agenda.
 *
 * Nada é pago automaticamente: a ferramenta apenas lê e organiza.
 */

// ===================== CONFIGURAÇÕES =====================
const CONFIG = {
  DIAS_ALERTA: 5,              // avisar contas que vencem nos próximos X dias
  BUSCAR_ULTIMOS_DIAS: 45,     // olhar e-mails recebidos nos últimos X dias
  HORARIOS_BUSCA: [7, 12, 18], // horários das buscas automáticas (0-23); pode pôr quantos quiser
  HORA_ALERTA: 8,              // hora do e-mail de alerta diário (0-23)
  CRIAR_EVENTO_AGENDA: true,   // criar lembrete no Google Agenda no dia do vencimento
  LER_PDF_ANEXO: true,         // ler o PDF quando o corpo do e-mail não tiver os dados

  // SÓ NAS CONTAS EXTRAS: cole entre as aspas o link da planilha principal.
  // Na conta principal (onde a planilha foi criada), deixe vazio.
  PLANILHA_PRINCIPAL: ''
};

// Fornecedores iniciais (depois você edita direto na aba "Fornecedores")
const FORNECEDORES_INICIAIS = []; // cada pessoa cadastra as próprias empresas no app

const ABA_CONTAS = 'Contas';
const ABA_FORN = 'Fornecedores';
const CAB_CONTAS = ['Fornecedor', 'Recebido em', 'Vencimento', 'Valor', 'Linha digitável',
                    'Status', 'Assunto', 'Link do e-mail', 'Origem dos dados', 'ID',
                    'Pix copia e cola', 'Observação', 'Conta que recebeu'];
const CAB_FORN = ['Nome', 'E-mail do remetente', 'Palavra no assunto (opcional)',
                  'Dia fixo de vencimento (opcional)'];

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 Boletos')
    .addItem('1. Instalar (autorizar e configurar)', 'configurar')
    .addItem('2. Criar ou alterar meu login', 'abrirCriarLogin')
    .addItem('3. Gerar meu cadastro para o app', 'mostrarCadastroApp')
    .addSeparator()
    .addItem('Buscar boletos agora', 'buscarBoletos')
    .addItem('Enviar alerta agora', 'enviarAlerta')
    .addItem('Sair de todos os aparelhos', 'sairDeTodosOsAparelhos')
    .addToUi();
}

// ===================== CONFIGURAÇÃO INICIAL =====================
function configurar() {
  const ss = planilha();

  // Aba Fornecedores
  let forn = ss.getSheetByName(ABA_FORN);
  if (!forn) {
    forn = ss.insertSheet(ABA_FORN);
    forn.getRange(1, 1, 1, CAB_FORN.length).setValues([CAB_FORN]).setFontWeight('bold');
    if (FORNECEDORES_INICIAIS.length)
      forn.getRange(2, 1, FORNECEDORES_INICIAIS.length, 4).setValues(FORNECEDORES_INICIAIS);
    forn.setFrozenRows(1);
    forn.autoResizeColumns(1, 4);
  }

  // Aba Contas
  let contas = ss.getSheetByName(ABA_CONTAS);
  if (!contas) contas = ss.insertSheet(ABA_CONTAS, 0);
  contas.getRange(1, 1, 1, CAB_CONTAS.length).setValues([CAB_CONTAS])
        .setFontWeight('bold').setBackground('#1f3b5c').setFontColor('#ffffff');
  contas.setFrozenRows(1);
  contas.getRange('B2:C').setNumberFormat('dd/mm/yyyy');
  contas.getRange('D2:D').setNumberFormat('"R$" #,##0.00');
  contas.getRange('E2:E').setNumberFormat('@');
  contas.getRange('F2:F').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['A pagar', 'Pago'], true).build());
  contas.hideColumns(10); // coluna ID (controle interno)

  // Cores: vencido = vermelho, vence em breve = amarelo, pago = verde
  const area = contas.getRange('A2:I');
  contas.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="Pago"')
      .setBackground('#d9f2d9').setRanges([area]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F2="A pagar";$C2<>"";$C2<TODAY())')
      .setBackground('#f8d0d0').setRanges([area]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F2="A pagar";$C2<>"";$C2-TODAY()<=' + CONFIG.DIAS_ALERTA + ')')
      .setBackground('#fff2c2').setRanges([area]).build()
  ]);

  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 2) ss.deleteSheet(padrao);

  // Agendamentos automáticos (recria para não duplicar)
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  CONFIG.HORARIOS_BUSCA.forEach(h =>
    ScriptApp.newTrigger('buscarBoletos').timeBased().everyDays(1).atHour(h).create());
  ScriptApp.newTrigger('enviarAlerta').timeBased().everyDays(1).atHour(CONFIG.HORA_ALERTA).create();

  buscarBoletos();
  try { SpreadsheetApp.getActive().toast(
    'Pronto! Buscas às ' + CONFIG.HORARIOS_BUSCA.join('h, ') + 'h e alerta às ' + CONFIG.HORA_ALERTA + 'h.',
    'Boletos', 8); } catch (e) {}
}

// Planilha usada pela ferramenta (funciona no menu, nos agendamentos e no app web)
function planilha() {
  if (CONFIG.PLANILHA_PRINCIPAL) {
    const ref = String(CONFIG.PLANILHA_PRINCIPAL).trim();
    return /^https?:/.test(ref) ? SpreadsheetApp.openByUrl(ref) : SpreadsheetApp.openById(ref);
  }
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) { props.setProperty('PLANILHA_ID', ss.getId()); return ss; }
  const id = props.getProperty('PLANILHA_ID');
  if (!id) throw new Error('Abra a planilha e rode "configurar" uma vez.');
  return SpreadsheetApp.openById(id);
}

// ===================== APP WEB (visual estilo aplicativo) =====================
function doGet(e) {
  // A página antiga só abre para o dono. Na implantação pública (usada pelo app do GitHub)
  // quem acessa sem login cai aqui e não vê nada.
  if (!ehDono()) {
    return ContentService.createTextOutput('Acesso restrito. Use o app Meus Boletos.')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  const p = (e && e.parameter) || {};
  let msg = '';
  let pdfs = null;

  // Evita repetir uma ação se a página for recarregada (F5) com o mesmo endereço
  const cache = CacheService.getUserCache();
  if (p.acao && p.n) {
    if (cache.get('acao_' + p.n)) p.acao = '';
    else cache.put('acao_' + p.n, '1', 21600);
  }

  try {
    if (p.acao === 'adicionar') {
      adicionarManual({ fornecedor: p.forn, obs: p.obs, vencimento: p.venc,
                        valor: p.val === '' || p.val == null ? null : Number(p.val), linha: p.linha, pix: p.pix });
      msg = 'Boleto adicionado';
    } else if (p.acao === 'pix' && p.id) {
      atualizarConta(p.id, { pix: p.pix });
      msg = 'Pix salvo';
    } else if (p.acao === 'addforn') {
      msg = adicionarFornecedor({ nome: p.nome, email: p.email, assunto: p.assunto, dia: p.dia }).msg;
    } else if (p.acao === 'delforn' && p.email) {
      msg = removerFornecedor(p.email, p.assunto).msg;
    } else if (p.acao === 'excluir' && p.id) {
      excluirConta(p.id);
      msg = 'Conta excluída';
    } else if (p.acao === 'lerpix' && p.id) {
      pdfs = { id: p.id, arquivos: getPdfs(p.id) };
    } else if (p.acao === 'status' && p.id) {
      atualizarConta(p.id, { status: p.valor === 'Pago' ? 'Pago' : 'A pagar' });
      msg = p.valor === 'Pago' ? 'Marcado como pago' : 'Voltou para "a pagar"';
    } else if (p.acao === 'editar' && p.id) {
      const campos = { vencimento: p.venc || '', valor: p.val === '' || p.val == null ? null : Number(p.val) };
      if (p.obs !== undefined) campos.obs = p.obs;
      atualizarConta(p.id, campos);
      msg = 'Alterações salvas';
    } else if (p.acao === 'buscar') {
      const n = buscarBoletos();
      msg = n > 0 ? n + ' boleto(s) novo(s)' : 'Nenhum boleto novo';
    }
  } catch (err) { msg = 'Erro: ' + err.message; }

  let dados = [], erro = '';
  try { dados = getContas(); } catch (err) { erro = err.message; }

  const json = v => JSON.stringify(v).replace(/</g, '\\u003c');
  const t = HtmlService.createTemplateFromFile('Pagina');
  t.dadosJson = json(dados);
  t.erroJson = json(erro);
  t.msgJson = json(msg);
  t.urlJson = json(ScriptApp.getService().getUrl() || '');
  t.pdfsJson = json(pdfs);
  let forns = [];
  try { forns = getFornecedores(); } catch (err) {}
  t.fornJson = json(forns);
  let caixas = [];
  try { caixas = getCaixas(); } catch (err) {}
  t.caixasJson = json(caixas);
  return t.evaluate()
    .setTitle('Meus Boletos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getContas() {
  const contas = planilha().getSheetByName(ABA_CONTAS);
  if (!contas || contas.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '';
  const emailDe = {};
  getFornecedores().forEach(f => { if (!emailDe[f.nome]) emailDe[f.nome] = f.email; });
  return contas.getRange(2, 1, contas.getLastRow() - 1, CAB_CONTAS.length).getValues()
    .filter(r => r[9])
    .map(r => ({
      fornecedor: String(r[0]), recebido: fmt(r[1]), vencimento: fmt(r[2]),
      valor: typeof r[3] === 'number' ? r[3] : null, linha: String(r[4] || ''),
      status: String(r[5] || 'A pagar'), assunto: String(r[6] || ''), link: String(r[7] || ''),
      origem: String(r[8] || ''), id: String(r[9]),
      pix: String(r[10] || ''), obs: String(r[11] || ''),
      remetente: /^manual-/.test(String(r[9])) ? '' : (emailDe[String(r[0])] || ''),
      conta: String(r[12] || '')
    }));
}

function garantirCabecalho(contas) {
  const atual = contas.getRange(1, 1, 1, CAB_CONTAS.length).getValues()[0];
  if (atual.join('|') !== CAB_CONTAS.join('|')) {
    contas.getRange(1, 1, 1, CAB_CONTAS.length).setValues([CAB_CONTAS])
          .setFontWeight('bold').setBackground('#1f3b5c').setFontColor('#ffffff');
    contas.getRange('K2:M').setNumberFormat('@');
  }
}

function linhaDaConta(contas, id) {
  if (contas.getLastRow() < 2) return -1;
  const ids = contas.getRange(2, 10, contas.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

// ===================== VÁRIAS CONTAS DE E-MAIL =====================
// Cada conta Gmail extra roda uma cópia deste script, que busca os boletos
// na própria caixa e grava na mesma planilha da conta principal.
const ABA_CAIXAS = 'Contas conectadas';

function minhaConta() {
  try { return Session.getEffectiveUser().getEmail().toLowerCase(); } catch (e) { return ''; }
}

// Rodar UMA vez em cada conta extra (depois de colar o link em PLANILHA_PRINCIPAL)
function configurarContaExtra() {
  if (!CONFIG.PLANILHA_PRINCIPAL) throw new Error('Cole o link da planilha principal em PLANILHA_PRINCIPAL, no topo do código.');
  const ss = planilha(); // testa o acesso à planilha
  if (!ss.getSheetByName(ABA_CONTAS)) throw new Error('A planilha informada não tem a aba "Contas". Confira o link.');
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // Contas extras buscam meia hora antes da principal, para o alerta já incluir tudo
  CONFIG.HORARIOS_BUSCA.forEach(h =>
    ScriptApp.newTrigger('buscarBoletos').timeBased().everyDays(1)
      .atHour(Math.max(0, h - 1)).nearMinute(30).create());
  const n = buscarBoletos();
  console.log('Conta ' + minhaConta() + ' conectada. ' + n + ' boleto(s) encontrado(s). Buscas automáticas por volta de ' +
              CONFIG.HORARIOS_BUSCA.map(h => Math.max(0, h - 1) + 'h30').join(', ') + '.');
}

function registrarCaixa(encontrados) {
  const ss = planilha();
  let aba = ss.getSheetByName(ABA_CAIXAS);
  if (!aba) {
    aba = ss.insertSheet(ABA_CAIXAS);
    aba.getRange(1, 1, 1, 4).setValues([['Conta', 'Tipo', 'Última busca', 'Boletos novos na última busca']])
       .setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  const conta = minhaConta();
  const tipo = CONFIG.PLANILHA_PRINCIPAL ? 'extra' : 'principal';
  const n = aba.getLastRow() - 1;
  const lista = n > 0 ? aba.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).toLowerCase()) : [];
  const i = lista.indexOf(conta);
  const linha = [conta, tipo, new Date(), encontrados];
  if (i >= 0) aba.getRange(i + 2, 1, 1, 4).setValues([linha]); else aba.appendRow(linha);
}

function getCaixas() {
  const aba = planilha().getSheetByName(ABA_CAIXAS);
  if (!aba || aba.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone();
  return aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues().filter(r => r[0]).map(r => ({
    conta: String(r[0]), tipo: String(r[1]),
    ultima: r[2] instanceof Date ? Utilities.formatDate(r[2], tz, "dd/MM 'às' HH:mm") : '',
    novos: r[3]
  }));
}

// ===================== CONTA DE E-MAIL QUE RECEBEU =====================
// Descobre em qual caixa de e-mail sua o boleto chegou.
// 1) Se veio por encaminhamento automático do Gmail, usa a conta que encaminhou.
// 2) Senão, usa o destinatário do e-mail (a conta principal, se estiver entre eles).
function contaDestino(msg) {
  const emails = txt => (String(txt || '').match(/[\w.+-]+@[\w-]+(\.[\w-]+)+/g) || []).map(e => e.toLowerCase());
  let principal = '';
  try { principal = Session.getEffectiveUser().getEmail().toLowerCase(); } catch (e) {}
  try {
    const fwd = emails(msg.getHeader('X-Forwarded-For')).filter(e => e !== principal);
    if (fwd.length) return fwd[0];
  } catch (e) {}
  const dest = emails(msg.getTo()).concat(emails(msg.getCc()));
  if (principal && dest.indexOf(principal) >= 0) return principal;
  return dest[0] || principal;
}

// Completa a coluna "Conta que recebeu" dos boletos lidos antes desta versão
function preencherContaAntigas(contas) {
  const n = contas.getLastRow() - 1;
  if (n < 1) return;
  const vals = contas.getRange(2, 10, n, 4).getValues(); // ID .. Conta
  let feitos = 0, tentativas = 0;
  vals.forEach((r, i) => {
    if (feitos >= 40 || tentativas >= 60 || r[3] || !r[0] || /^manual-/.test(String(r[0]))) return;
    tentativas++;
    try {
      contas.getRange(i + 2, 13).setValue(contaDestino(GmailApp.getMessageById(String(r[0]))));
      feitos++;
    } catch (e) {}
  });
}

// ===================== EMPRESAS (fornecedores) =====================
function getFornecedores() {
  const forn = planilha().getSheetByName(ABA_FORN);
  if (!forn || forn.getLastRow() < 2) return [];
  return forn.getRange(2, 1, forn.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ nome: String(r[0]), email: String(r[1]).trim(), assunto: String(r[2] || ''),
                 dia: r[3] === '' ? '' : String(r[3]) }));
}

function estado(msg) {
  let caixas = [];
  try { caixas = getCaixas(); } catch (e) {}
  let sugestoes = [];
  try { sugestoes = getSugestoes(); } catch (e) {}
  return { contas: getContas(), fornecedores: getFornecedores(), caixas: caixas, sugestoes: sugestoes, msg: msg || '' };
}

function adicionarFornecedor(f) {
  const nome = String(f.nome || '').trim();
  const email = String(f.email || '').trim().toLowerCase();
  const assunto = String(f.assunto || '').trim();
  const dia = String(f.dia || '').trim();
  if (!nome) throw new Error('Informe o nome da empresa.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('E-mail do remetente inválido.');
  if (dia && !(Number(dia) >= 1 && Number(dia) <= 31)) throw new Error('Dia de vencimento deve ser de 1 a 31.');

  const forn = planilha().getSheetByName(ABA_FORN);
  const existe = getFornecedores().some(x => x.email.toLowerCase() === email && x.assunto.toLowerCase() === assunto.toLowerCase());
  if (!existe) forn.appendRow([nome, email, assunto, dia ? Number(dia) : '']);
  const novos = buscarBoletos();
  return estado((existe ? 'Essa empresa já estava cadastrada' : nome + ' cadastrada') +
                ' · ' + (novos > 0 ? novos + ' boleto(s) encontrado(s)' : 'nenhum boleto nos últimos ' + CONFIG.BUSCAR_ULTIMOS_DIAS + ' dias'));
}

function removerFornecedor(email, assunto) {
  const forn = planilha().getSheetByName(ABA_FORN);
  const alvoE = String(email || '').trim().toLowerCase(), alvoA = String(assunto || '').trim().toLowerCase();
  for (let l = forn.getLastRow(); l >= 2; l--) {
    const r = forn.getRange(l, 1, 1, 3).getValues()[0];
    if (String(r[1]).trim().toLowerCase() === alvoE && String(r[2] || '').trim().toLowerCase() === alvoA) forn.deleteRow(l);
  }
  return estado('Empresa removida (os boletos que já estão na lista continuam)');
}

// Inclusão manual (boleto que não chegou por e-mail)
function adicionarManual(c) {
  const contas = planilha().getSheetByName(ABA_CONTAS);
  garantirCabecalho(contas);
  const nome = String(c.fornecedor || '').trim();
  if (!nome) throw new Error('Informe a que o boleto se refere.');

  const digitos = String(c.linha || '').replace(/\D/g, '');
  let linha = String(c.linha || '').trim();
  let dadosLinha = { valor: null, vencimento: null };
  if (digitos.length === 44 && digitos[0] !== '8') {
    const l = barrasParaLinha(digitos);
    if (l) { linha = formatarLinha47(l); dadosLinha = extrairDados(l); }
  } else if (digitos.length >= 47) {
    const d = extrairDados(digitos);
    if (d.linha) { linha = d.linha; dadosLinha = d; }
  }

  const p = String(c.vencimento || '').split('-');
  const venc = p.length === 3 ? new Date(+p[0], +p[1] - 1, +p[2]) : dadosLinha.vencimento;
  const valor = c.valor != null && !isNaN(c.valor) ? Number(c.valor) : dadosLinha.valor;
  const pix = pixValido(String(c.pix || '').trim()) ? String(c.pix).trim() : '';
  const id = 'manual-' + Date.now();

  contas.appendRow([nome, new Date(), venc || '', valor || '', linha, 'A pagar', 'Adicionado no app', '',
                    'adicionado manualmente', id, pix, String(c.obs || ''), '']);
  const n = contas.getLastRow() - 1;
  if (n > 1) contas.getRange(2, 1, n, CAB_CONTAS.length).sort({ column: 3, ascending: true });
  if (CONFIG.CRIAR_EVENTO_AGENDA && venc) criarLembrete(nome, { vencimento: venc, valor: valor, linha: linha });
  return getContas();
}

function excluirConta(id) {
  const contas = planilha().getSheetByName(ABA_CONTAS);
  const l = linhaDaConta(contas, id);
  if (l < 0) throw new Error('Conta não encontrada.');
  contas.deleteRow(l);
  return getContas();
}

// PDFs anexos do e-mail, para o app ler o QR Code do Pix no navegador
function getPdfs(id) {
  if (!id || /^manual-/.test(id)) return [];
  let msg;
  try { msg = GmailApp.getMessageById(id); }
  catch (e) { throw new Error('Esse boleto chegou em outra conta de e-mail. Abra o e-mail nela e use "Colar Pix".'); }
  const peso = n => /boleto/i.test(n) ? 0 : /fatura/i.test(n) ? 1 : 2;
  return msg.getAttachments()
    .filter(a => a.getContentType() === 'application/pdf' || /\.pdf$/i.test(a.getName()))
    .sort((a, b) => peso(a.getName()) - peso(b.getName()))
    .slice(0, 3)
    .map(a => ({ nome: a.getName(), b64: Utilities.base64Encode(a.getBytes()) }));
}

function atualizarConta(id, campos) {
  const contas = planilha().getSheetByName(ABA_CONTAS);
  garantirCabecalho(contas);
  const linha = linhaDaConta(contas, id);
  if (linha < 0) throw new Error('Conta não encontrada.');
  if (campos.pix !== undefined) {
    const pix = String(campos.pix || '').trim();
    if (pix && !pixValido(pix)) throw new Error('Código Pix inválido.');
    contas.getRange(linha, 11).setValue(pix);
  }
  if (campos.obs !== undefined) contas.getRange(linha, 12).setValue(String(campos.obs || ''));
  if (campos.status) contas.getRange(linha, 6).setValue(campos.status);
  if (campos.vencimento !== undefined) {
    const p = String(campos.vencimento).split('-');
    contas.getRange(linha, 3).setValue(p.length === 3 ? new Date(+p[0], +p[1] - 1, +p[2]) : '');
  }
  if (campos.valor !== undefined) contas.getRange(linha, 4).setValue(campos.valor === null ? '' : Number(campos.valor));
  return getContas();
}

// ===================== API PARA O APP DO GITHUB =====================
// O app (página no GitHub Pages) conversa com este script por POST.
// 1) {acao:'login', email, senha}  -> devolve um token de sessão
// 2) {token, fn, args}             -> executa uma das funções permitidas
const FUNCOES_API = {
  estado:              () => estado(''),
  buscarAgora:         () => { const n = buscarBoletos(); return estado(n > 0 ? n + ' boleto(s) novo(s)' : 'Nenhum boleto novo'); },
  atualizarConta:      (id, campos) => atualizarConta(id, campos),
  adicionarManual:     (c) => adicionarManual(c),
  excluirConta:        (id) => excluirConta(id),
  getPdfs:             (id) => getPdfs(id),
  adicionarFornecedor: (f) => adicionarFornecedor(f),
  removerFornecedor:   (email, assunto) => removerFornecedor(email, assunto),
  aceitarSugestao:     (id, nome) => aceitarSugestao(id, nome),
  ignorarSugestao:     (id) => ignorarSugestao(id)
};
const DIAS_SESSAO = 180;
const MAX_TENTATIVAS = 5;

function doPost(e) {
  let resposta;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.acao === 'login') {
      resposta = { ok: true, dados: fazerLogin(req.email, req.senha, req.aparelho) };
    } else {
      const sessao = validarSessao(req.token);
      if (!sessao) {
        resposta = { ok: false, erro: 'sessao', mensagem: 'Sua sessão expirou. Entre novamente.' };
      } else if (req.fn === 'sair') {
        encerrarSessao(req.token);
        resposta = { ok: true, dados: true };
      } else {
        const fn = FUNCOES_API[req.fn];
        if (!fn) throw new Error('Função não permitida.');
        resposta = { ok: true, dados: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
      }
    }
  } catch (err) {
    resposta = { ok: false, erro: err.codigo || 'falha', mensagem: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(resposta)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Login ----------
function erro(codigo, msg) { const e = new Error(msg); e.codigo = codigo; return e; }

function fazerLogin(email, senha, aparelho) {
  const props = PropertiesService.getScriptProperties();
  const cad = JSON.parse(props.getProperty('LOGIN') || 'null');
  if (!cad) throw erro('login', 'Este app ainda não tem login criado. Na planilha, use o menu 💰 Boletos → Criar ou alterar meu login.');

  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get('falhas_login') || 0);
  if (falhas >= MAX_TENTATIVAS) throw erro('bloqueado', 'Muitas tentativas erradas. Aguarde 15 minutos e tente de novo.');

  const okEmail = String(email || '').trim().toLowerCase() === cad.email;
  const okSenha = iguais(hashSenha(String(senha || ''), cad.sal), cad.hash);
  if (!okEmail || !okSenha) {
    cache.put('falhas_login', String(falhas + 1), 900);
    Utilities.sleep(1200);
    throw erro('login', 'E-mail ou senha incorretos.');
  }
  cache.remove('falhas_login');

  const token = gerarAleatorio(40);
  const sessoes = lerSessoes();
  sessoes[sha256(token)] = { exp: Date.now() + DIAS_SESSAO * 864e5, criada: Date.now(),
                             aparelho: String(aparelho || '').slice(0, 60) };
  gravarSessoes(sessoes);
  return { token: token, email: cad.email };
}

function validarSessao(token) {
  if (!token) return null;
  const sessoes = lerSessoes();
  const s = sessoes[sha256(String(token))];
  if (!s || s.exp < Date.now()) return null;
  return s;
}

function encerrarSessao(token) {
  const sessoes = lerSessoes();
  delete sessoes[sha256(String(token))];
  gravarSessoes(sessoes);
}

function lerSessoes() {
  const todas = JSON.parse(PropertiesService.getScriptProperties().getProperty('SESSOES') || '{}');
  const agora = Date.now();
  Object.keys(todas).forEach(k => { if (todas[k].exp < agora) delete todas[k]; });
  return todas;
}
function gravarSessoes(s) { PropertiesService.getScriptProperties().setProperty('SESSOES', JSON.stringify(s)); }

// Senha nunca é guardada: só um "hash" lento com sal, que não dá para desfazer
function hashSenha(senha, sal) {
  const chave = Utilities.newBlob(senha).getBytes();
  let h = Utilities.computeHmacSha256Signature(Utilities.newBlob('boletos:' + sal).getBytes(), chave);
  for (let i = 0; i < 3000; i++) h = Utilities.computeHmacSha256Signature(h, chave);
  return Utilities.base64Encode(h);
}
function sha256(txt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt));
}
function gerarAleatorio(n) {
  const letras = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  let bytes = [];
  while (bytes.length < n) bytes = bytes.concat(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + Utilities.getUuid() + Date.now() + Math.random()));
  return bytes.slice(0, n).map(b => letras[(b + 256) % letras.length]).join('');
}

// Comparação que não revela, pelo tempo, quantos caracteres estão certos
function iguais(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function ehDono() {
  try {
    const quem = Session.getActiveUser().getEmail();
    return !!quem && quem.toLowerCase() === Session.getEffectiveUser().getEmail().toLowerCase();
  } catch (e) { return false; }
}

// ---------- Criar / alterar o login (menu da planilha) ----------
function abrirCriarLogin() {
  const email = Session.getEffectiveUser().getEmail();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p>Seu e-mail de login: <b>' + email + '</b></p>' +
    '<label>Nova senha (mínimo 8 caracteres)<br><input id="s1" type="password" style="width:100%;padding:6px"></label><br><br>' +
    '<label>Repita a senha<br><input id="s2" type="password" style="width:100%;padding:6px"></label>' +
    '<p id="m" style="color:#c0392b;min-height:1.2em"></p>' +
    '<button id="b" style="padding:8px 16px">Salvar login</button>' +
    '<script>document.getElementById("b").onclick=function(){var a=s1.value,b=s2.value,m=document.getElementById("m");' +
    'if(a.length<8){m.textContent="A senha precisa ter pelo menos 8 caracteres.";return}' +
    'if(a!==b){m.textContent="As senhas não são iguais.";return}' +
    'this.disabled=true;m.style.color="#333";m.textContent="Salvando…";' +
    'google.script.run.withSuccessHandler(function(){m.style.color="#1b7f6b";m.textContent="Login salvo! Pode fechar esta janela.";})' +
    '.withFailureHandler(function(e){m.style.color="#c0392b";m.textContent=e.message;document.getElementById("b").disabled=false;}).salvarLogin(a);};</script></div>'
  ).setWidth(360).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Criar ou alterar meu login');
}

function salvarLogin(senha) {
  senha = String(senha || '');
  if (senha.length < 8) throw new Error('A senha precisa ter pelo menos 8 caracteres.');
  const sal = gerarAleatorio(24);
  PropertiesService.getScriptProperties().setProperty('LOGIN', JSON.stringify({
    email: Session.getEffectiveUser().getEmail().toLowerCase(), sal: sal, hash: hashSenha(senha, sal)
  }));
  PropertiesService.getScriptProperties().deleteProperty('SESSOES'); // troca de senha desconecta todos os aparelhos
  PropertiesService.getScriptProperties().deleteProperty('SENHA_APP'); // senha antiga do app deixa de valer
  return true;
}

// Passo 3 da instalação: gera a mensagem que o novo usuário manda para quem administra o app
function mostrarCadastroApp() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LOGIN')) { ui.alert('Antes, faça o passo 2: menu 💰 Boletos → 2. Criar ou alterar meu login.'); return; }
  const url = ScriptApp.getService().getUrl();
  if (!url || !/\/exec$/.test(url)) {
    ui.alert('Ainda falta publicar o motor: no editor do script, clique em Implantar → Nova implantação → ' +
             'App da Web, "Executar como: Eu", "Quem pode acessar: Qualquer pessoa" → Implantar. Depois volte aqui.');
    return;
  }
  const email = Session.getEffectiveUser().getEmail().toLowerCase();
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  const msg = 'Cadastro Central dos Boletos (' + email + ')\n"' + hash + '": "' + url + '"';
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p>Copie a mensagem abaixo e envie para quem te convidou para o app (WhatsApp, e-mail…).</p>' +
    '<textarea id="t" readonly style="width:100%;height:110px;font-family:monospace;font-size:12px">' +
    msg.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</textarea>' +
    '<p><button onclick="t.select();document.execCommand(\'copy\');this.textContent=\'Copiado!\'">Copiar mensagem</button></p>' +
    '<p style="color:#555">Ela não contém sua senha. Quando seu acesso for liberado, entre em ' +
    '<b>nobregadelivery.github.io/centraldosboletos</b> com <b>' + email + '</b> e a senha que você criou.</p></div>'
  ).setWidth(440).setHeight(330);
  ui.showModalDialog(html, 'Seu cadastro para o app');
}

function sairDeTodosOsAparelhos() {
  PropertiesService.getScriptProperties().deleteProperty('SESSOES');
  try { SpreadsheetApp.getUi().alert('Pronto: todos os aparelhos foram desconectados.'); } catch (e) {}
}

function buscarAgora() {
  buscarBoletos();
  return getContas();
}

// ===================== BUSCA DE BOLETOS =====================
function buscarBoletos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) return 0;
  try { return buscarBoletosSemTrava(); } finally { lock.releaseLock(); }
}

function buscarBoletosSemTrava() {
  const ss = planilha();
  const contas = ss.getSheetByName(ABA_CONTAS);
  const forn = ss.getSheetByName(ABA_FORN);
  if (!contas || !forn) throw new Error('Rode "Configurar (rodar 1ª vez)" no menu 💰 Boletos.');
  garantirCabecalho(contas);

  const ultima = contas.getLastRow();
  const idsJaLidos = new Set(ultima > 1
    ? contas.getRange(2, 10, ultima - 1, 1).getValues().map(r => String(r[0])) : []);

  const fornecedores = forn.getLastRow() > 1
    ? forn.getRange(2, 1, forn.getLastRow() - 1, 4).getValues().filter(r => r[0] && r[1]) : [];

  preencherContaAntigas(contas);

  const novas = [];
  fornecedores.forEach(([nome, remetente, assunto, diaFixo]) => {
    let q = 'from:' + String(remetente).trim() + ' newer_than:' + CONFIG.BUSCAR_ULTIMOS_DIAS + 'd';
    if (assunto) q += ' subject:"' + String(assunto).replace(/"/g, '') + '"';

    GmailApp.search(q, 0, 20).forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (idsJaLidos.has(id)) return;
        if (String(msg.getFrom()).toLowerCase().indexOf(String(remetente).trim().toLowerCase()) < 0) return;

        // Aviso de DDA ("novos boletos emitidos no seu CNPJ/CPF"): só traz quem emitiu
        const aviso = avisoDDA(msg.getSubject(), msg.getPlainBody() || '', msg.getFrom());
        if (aviso) {
          const linkAviso = 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(minhaConta()) +
                            '#all/' + msg.getThread().getId();
          novas.push([aviso.emissor, msg.getDate(), '', '', '', 'A pagar', msg.getSubject(), linkAviso,
                      'aviso DDA (' + aviso.banco + ')', id, '',
                      'Confira valor e vencimento no app do ' + aviso.banco, contaDestino(msg)]);
          idsJaLidos.add(id);
          return;
        }

        let dados = extrairDados(msg.getPlainBody() || '');
        let origem = 'corpo do e-mail';

        if (CONFIG.LER_PDF_ANEXO && (!dados.vencimento || !dados.valor || !dados.pix)) {
          const doPdf = lerPdfs(msg);
          if (doPdf) {
            if ((!dados.vencimento || !dados.valor) && (doPdf.vencimento || doPdf.valor)) origem = 'PDF anexo';
            dados = {
              valor: dados.valor || doPdf.valor,
              vencimento: dados.vencimento || doPdf.vencimento,
              linha: dados.linha || doPdf.linha,
              pix: dados.pix || doPdf.pix
            };
          }
        }

        if (!dados.vencimento && diaFixo) {
          dados.vencimento = proximoDia(msg.getDate(), Number(diaFixo));
          origem += ' + dia fixo';
        }
        if (!dados.vencimento && !dados.valor) origem = 'não encontrado — confira o e-mail';

        const link = 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(minhaConta()) +
                     '#all/' + msg.getThread().getId();
        novas.push([nome, msg.getDate(), dados.vencimento || '', dados.valor || '',
                    dados.linha || '', 'A pagar', msg.getSubject(), link, origem, id,
                    dados.pix || '', '', contaDestino(msg)]);
        idsJaLidos.add(id);

        if (CONFIG.CRIAR_EVENTO_AGENDA && dados.vencimento) criarLembrete(nome, dados);
      });
    });
  });

  if (novas.length) {
    // appendRow linha a linha: seguro mesmo se outra conta estiver gravando ao mesmo tempo
    novas.forEach(r => contas.appendRow(r));
    const n = contas.getLastRow() - 1;
    contas.getRange(2, 1, n, CAB_CONTAS.length).sort({ column: 3, ascending: true });
  }
  registrarCaixa(novas.length);
  try { buscarSugestoesSeHoje(); } catch (e) { console.warn('Sugestões: ' + e); }
  try { ss.toast(novas.length + ' boleto(s) novo(s) encontrado(s).', 'Boletos', 5); } catch (e) {}
  return novas.length;
}

// ===================== SUGESTÕES DE BOLETO PARA CADASTRO =====================
// Uma vez por dia procura e-mails com cara de boleto vindos de remetentes que ainda
// não são empresas cadastradas. Ficam na aba "Sugestões" até a pessoa adicionar ou ignorar.
const ABA_SUG = 'Sugestões';
const CAB_SUG = ['ID', 'E-mail do remetente', 'Nome', 'Assunto', 'Recebido em', 'Link', 'Conta', 'Status', 'Sinal'];

function abaSugestoes() {
  const ss = planilha();
  let aba = ss.getSheetByName(ABA_SUG);
  if (!aba) {
    aba = ss.insertSheet(ABA_SUG);
    aba.getRange(1, 1, 1, CAB_SUG.length).setValues([CAB_SUG]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function buscarSugestoesSeHoje() {
  const props = PropertiesService.getScriptProperties();
  const ultima = Number(props.getProperty('SUG_ULTIMA') || 0);
  if (Date.now() - ultima < 20 * 3600 * 1000) return 0;
  const dias = ultima ? 3 : 30; // na primeira vez olha 30 dias para trás
  const n = buscarSugestoes(dias);
  props.setProperty('SUG_ULTIMA', String(Date.now()));
  return n;
}

function buscarSugestoes(dias) {
  const aba = abaSugestoes();
  const conhecidos = new Set(getFornecedores().map(f => f.email.toLowerCase()));
  const n = aba.getLastRow() - 1;
  const linhas = n > 0 ? aba.getRange(2, 1, n, CAB_SUG.length).getValues() : [];
  const jaVistos = new Set(linhas.map(r => String(r[1]).toLowerCase())); // um remetente só uma vez
  const eu = minhaConta();

  const q = 'newer_than:' + dias + 'd -category:promotions -category:social -in:chats -from:me ' +
            '(subject:(boleto OR fatura OR cobrança OR "segunda via" OR "2ª via") OR filename:boleto OR filename:fatura OR "linha digitável")';
  const novas = [];
  GmailApp.search(q, 0, 40).forEach(th => {
    th.getMessages().forEach(msg => {
      const de = String(msg.getFrom());
      const m = de.match(/[\w.+-]+@[\w-]+(\.[\w-]+)+/);
      if (!m) return;
      const email = m[0].toLowerCase();
      if (email === eu || conhecidos.has(email) || jaVistos.has(email)) return;
      const sinal = sinalDeBoleto(msg);
      if (!sinal) return;
      const nome = (de.replace(/<.*>/, '').replace(/"/g, '').trim()) || email.split('@')[1];
      novas.push([msg.getId(), email, nome.slice(0, 60), msg.getSubject().slice(0, 120), msg.getDate(),
                  'https://mail.google.com/mail/?authuser=' + encodeURIComponent(eu) + '#all/' + th.getId(),
                  contaDestino(msg), 'nova', sinal]);
      jaVistos.add(email);
    });
  });
  novas.forEach(r => aba.appendRow(r));
  return novas.length;
}

// Confere se o e-mail tem cara de boleto e diz por quê
function sinalDeBoleto(msg) {
  const assunto = String(msg.getSubject() || '');
  let corpo = '';
  try { corpo = String(msg.getPlainBody() || '').slice(0, 20000); } catch (e) {}
  const d = extrairDados(corpo);
  if (d.linha) return 'código de barras no e-mail';
  let pdfs = [];
  try { pdfs = msg.getAttachments().filter(a => /\.pdf$/i.test(a.getName()) || a.getContentType() === 'application/pdf'); } catch (e) {}
  if (pdfs.some(a => /boleto|fatura|cobran/i.test(a.getName()))) return 'PDF de boleto anexo';
  if (/boleto|fatura|cobran[cç]a|segunda via|2ª via/i.test(assunto) && (pdfs.length || d.pix || (d.valor && d.vencimento)))
    return pdfs.length ? 'assunto de cobrança + PDF anexo' : 'assunto de cobrança com valor e vencimento';
  return null;
}

function getSugestoes() {
  const aba = planilha().getSheetByName(ABA_SUG);
  if (!aba || aba.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone();
  return aba.getRange(2, 1, aba.getLastRow() - 1, CAB_SUG.length).getValues()
    .filter(r => r[0] && r[7] === 'nova')
    .map(r => ({ id: String(r[0]), email: String(r[1]), nome: String(r[2]), assunto: String(r[3]),
                 recebido: r[4] instanceof Date ? Utilities.formatDate(r[4], tz, 'yyyy-MM-dd') : '',
                 link: String(r[5]), conta: String(r[6]), sinal: String(r[8]) }))
    .sort((a, b) => a.recebido < b.recebido ? 1 : -1);
}

function marcarSugestao(id, status) {
  const aba = planilha().getSheetByName(ABA_SUG);
  if (!aba || aba.getLastRow() < 2) throw new Error('Sugestão não encontrada.');
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i < 0) throw new Error('Sugestão não encontrada.');
  aba.getRange(i + 2, 8).setValue(status);
  return aba.getRange(i + 2, 1, 1, CAB_SUG.length).getValues()[0];
}

function aceitarSugestao(id, nome) {
  const r = marcarSugestao(id, 'adicionada');
  return adicionarFornecedor({ nome: String(nome || r[2]).trim() || String(r[1]), email: String(r[1]), assunto: '', dia: '' });
}

function ignorarSugestao(id) {
  marcarSugestao(id, 'ignorada');
  return estado('Sugestão ignorada');
}

// ===================== AVISOS DE DDA =====================
// Ex.: Nubank "Novos boletos emitido no seu CNPJ" -> "...emitidos por BAHIANA DISTRIBUIDO."
function avisoDDA(assunto, corpo, de) {
  const t = String(corpo);
  if (!/novos boletos/i.test(String(assunto)) && !/identificou novos boletos/i.test(t)) return null;
  const m = t.match(/emitidos? por\s+(.+?)\.?\s*(?:\n|$)/i);
  const emissor = m ? m[1].trim().slice(0, 80) : 'Boleto novo no seu CNPJ/CPF';
  const banco = /nubank/i.test(t + ' ' + assunto + ' ' + (de || '')) ? 'Nubank' : 'banco';
  return { emissor: emissor, banco: banco };
}

// ===================== LEITURA DE PDF =====================
// Converte o PDF em Google Docs (OCR do próprio Google), lê o texto e apaga a cópia.
function lerPdfs(msg) {
  const pdfs = msg.getAttachments().filter(a =>
    a.getContentType() === 'application/pdf' || /\.pdf$/i.test(a.getName()));
  // Prioriza arquivos com "boleto" no nome, depois "fatura"
  const peso = n => /boleto/i.test(n) ? 0 : /fatura/i.test(n) ? 1 : 2;
  pdfs.sort((a, b) => peso(a.getName()) - peso(b.getName()));

  const r = { valor: null, vencimento: null, linha: null, pix: null };
  let achou = false;
  for (const pdf of pdfs) {
    let docId = null;
    try {
      const arq = Drive.Files.create(
        { name: 'tmp_boleto_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS },
        pdf.copyBlob(), { ocrLanguage: 'pt' });
      docId = arq.id;
      const texto = DocumentApp.openById(docId).getBody().getText();
      const d = extrairDados(texto);
      // Linha digitável e vencimento vêm juntos do mesmo PDF; Pix pode estar em outro
      if (!r.linha && !r.vencimento && (d.linha || d.vencimento)) {
        r.linha = d.linha; r.vencimento = d.vencimento; r.valor = d.valor || r.valor;
      } else if (!r.valor && d.valor) r.valor = d.valor;
      if (!r.pix && d.pix) r.pix = d.pix;
      achou = achou || !!(d.linha || d.vencimento || d.valor || d.pix);
      if (r.linha && r.pix) break;
    } catch (e) {
      console.warn('Falha ao ler PDF ' + pdf.getName() + ': ' + e);
    } finally {
      if (docId) try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
    }
  }
  return achou ? r : null;
}

// ===================== EXTRAÇÃO DOS DADOS =====================
function extrairDados(texto) {
  const t = String(texto).replace(/ /g, ' ');
  const r = { valor: null, vencimento: null, linha: null, pix: extrairPix(t) };

  // 1) Linha digitável de boleto bancário (47 dígitos) — só aceita se os
  //    dígitos verificadores conferirem (evita confundir com chave de nota fiscal etc.)
  const cand47 = [...t.matchAll(/\d{5}[.\s]?\d{5}[ \t]*\d{5}[.\s]?\d{6}[ \t]*\d{5}[.\s]?\d{6}[ \t]*\d[ \t]*\d{14}/g)]
    .map(m => m[0].replace(/\D/g, '')).find(linha47Valida);
  // 2) Linha de conta de consumo/concessionária (48 dígitos, começa com 8)
  const cand48 = [...t.matchAll(/8\d{10}[-\s]?\d[ \t]*\d{11}[-\s]?\d[ \t]*\d{11}[-\s]?\d[ \t]*\d{11}[-\s]?\d/g)]
    .map(m => m[0].replace(/\D/g, '')).find(linha48Valida);

  if (cand47) {
    const l = cand47;
    r.linha = formatarLinha47(l);
    const fator = Number(l.substr(33, 4));
    const valor = Number(l.substr(37, 10)) / 100;
    if (valor > 0) r.valor = valor;
    if (fator > 0) r.vencimento = dataPorFator(fator);
  } else if (cand48) {
    const l = cand48;
    r.linha = l.replace(/(\d{11})(\d)/g, '$1-$2 ').trim();
    const barras = [0, 12, 24, 36].map(i => l.substr(i, 11)).join('');
    if (barras[2] === '6' || barras[2] === '7') {
      const valor = Number(barras.substr(4, 11)) / 100;
      if (valor > 0) r.valor = valor;
    }
  }

  // Vencimento escrito no texto (tem prioridade se existir)
  const mv = t.match(/venc[a-zçã]*[^\d]{0,40}(\d{2})[\/.-](\d{2})[\/.-](\d{2,4})/i);
  if (mv) {
    const ano = mv[3].length === 2 ? 2000 + Number(mv[3]) : Number(mv[3]);
    const d = new Date(ano, Number(mv[2]) - 1, Number(mv[1]));
    if (!isNaN(d)) r.vencimento = d;
  }

  // Valor escrito no texto, se a linha digitável não trouxe
  if (!r.valor) {
    const mval = t.match(/(?:valor\s*(?:do\s*documento|cobrado|total|a\s*pagar)?|total\s*a\s*pagar)[^\dR]{0,30}(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})/i)
              || t.match(/R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})/);
    if (mval) r.valor = Number(mval[1].replace(/\./g, '').replace(',', '.'));
  }
  return r;
}

// ---- Validação pelos dígitos verificadores (padrão FEBRABAN) ----
function mod10(num) {
  let soma = 0, peso = 2;
  for (let i = num.length - 1; i >= 0; i--) {
    let p = Number(num[i]) * peso;
    if (p > 9) p = Math.floor(p / 10) + (p % 10);
    soma += p; peso = peso === 2 ? 1 : 2;
  }
  return (10 - (soma % 10)) % 10;
}

function mod11Boleto(num) {
  let soma = 0, peso = 2;
  for (let i = num.length - 1; i >= 0; i--) {
    soma += Number(num[i]) * peso; peso = peso === 9 ? 2 : peso + 1;
  }
  const dv = 11 - (soma % 11);
  return (dv === 0 || dv === 10 || dv === 11) ? 1 : dv;
}

function mod11Consumo(num) {
  let soma = 0, peso = 2;
  for (let i = num.length - 1; i >= 0; i--) {
    soma += Number(num[i]) * peso; peso = peso === 9 ? 2 : peso + 1;
  }
  const r = soma % 11;
  return (r === 0 || r === 1) ? 0 : 11 - r;
}

function linha47Valida(l) {
  if (!l || l.length !== 47) return false;
  if (mod10(l.substr(0, 9)) !== Number(l[9])) return false;
  if (mod10(l.substr(10, 10)) !== Number(l[20])) return false;
  if (mod10(l.substr(21, 10)) !== Number(l[31])) return false;
  const barras = l.substr(0, 4) + l.substr(33, 14) + l.substr(4, 5) + l.substr(10, 10) + l.substr(21, 10);
  return mod11Boleto(barras) === Number(l[32]);
}

function linha48Valida(l) {
  if (!l || l.length !== 48 || l[0] !== '8') return false;
  const tipo = l[2];
  const f = (tipo === '6' || tipo === '7') ? mod10 : (tipo === '8' || tipo === '9') ? mod11Consumo : null;
  if (!f) return false;
  for (let i = 0; i < 4; i++) {
    if (f(l.substr(i * 12, 11)) !== Number(l[i * 12 + 11])) return false;
  }
  return true;
}

// ---- Pix "copia e cola" (padrão EMV do Banco Central) ----
function crc16Pix(str) {
  let crc = 0xFFFF;
  for (let i = 0; i < str.length; i++) {
    crc ^= str.charCodeAt(i) << 8;
    for (let j = 0; j < 8; j++) crc = (crc & 0x8000) ? ((crc << 1) ^ 0x1021) & 0xFFFF : (crc << 1) & 0xFFFF;
  }
  return crc.toString(16).toUpperCase().padStart(4, '0');
}

function pixValido(s) {
  if (!s || s.length < 50 || s.slice(0, 6) !== '000201' || !/BR\.GOV\.BCB\.PIX/i.test(s)) return false;
  return s.slice(-8, -4) === '6304' && crc16Pix(s.slice(0, -4)) === s.slice(-4).toUpperCase();
}

// Procura um código Pix no texto (o PDF costuma quebrar o código em várias linhas)
function extrairPix(texto) {
  const variantes = [String(texto).replace(/\r?\n/g, ''), String(texto).replace(/\s+/g, '')];
  for (const v of variantes) {
    let ini = v.indexOf('000201');
    while (ini >= 0) {
      let fim = v.indexOf('6304', ini + 20);
      while (fim >= 0 && fim - ini < 700) {
        const cand = v.substr(ini, fim + 8 - ini);
        if (pixValido(cand)) return cand;
        fim = v.indexOf('6304', fim + 1);
      }
      ini = v.indexOf('000201', ini + 1);
    }
  }
  return null;
}

// Código de barras de 44 dígitos (lido por leitor) -> linha digitável de 47
function barrasParaLinha(b) {
  if (!b || b.length !== 44) return null;
  if (mod11Boleto(b.slice(0, 4) + b.slice(5)) !== Number(b[4])) return null;
  const c1 = b.substr(0, 4) + b.substr(19, 5), c2 = b.substr(24, 10), c3 = b.substr(34, 10);
  return c1 + mod10(c1) + c2 + mod10(c2) + c3 + mod10(c3) + b[4] + b.substr(5, 14);
}

// Fator de vencimento FEBRABAN. Reiniciou em 1000 em 22/02/2025.
function dataPorFator(fator) {
  const dia = 24 * 3600 * 1000;
  const antigo = new Date(new Date(1997, 9, 7).getTime() + fator * dia);
  const novo = new Date(new Date(2025, 1, 22).getTime() + (fator - 1000) * dia);
  const hoje = Date.now();
  return Math.abs(novo - hoje) < Math.abs(antigo - hoje) ? novo : antigo;
}

function formatarLinha47(l) {
  return l.substr(0, 5) + '.' + l.substr(5, 5) + ' ' + l.substr(10, 5) + '.' + l.substr(15, 6) + ' ' +
         l.substr(21, 5) + '.' + l.substr(26, 6) + ' ' + l.substr(32, 1) + ' ' + l.substr(33, 14);
}

function proximoDia(base, dia) {
  const d = new Date(base.getFullYear(), base.getMonth(), dia);
  if (d < base) d.setMonth(d.getMonth() + 1);
  return d;
}

// ===================== AGENDA =====================
function criarLembrete(nome, dados) {
  try {
    const valor = dados.valor ? ' — ' + moeda(dados.valor) : '';
    const ev = CalendarApp.getDefaultCalendar().createAllDayEvent(
      '💰 Vence boleto ' + nome + valor, dados.vencimento,
      { description: dados.linha ? 'Linha digitável: ' + dados.linha : '' });
    ev.removeAllReminders();
    ev.addPopupReminder(60 * 24 * 3); // 3 dias antes
    ev.addPopupReminder(60 * 12);     // na véspera
  } catch (e) { console.warn('Agenda: ' + e); }
}

// ===================== ALERTA DIÁRIO =====================
function enviarAlerta() {
  const contas = planilha().getSheetByName(ABA_CONTAS);
  if (!contas || contas.getLastRow() < 2) return;
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const limite = new Date(hoje.getTime() + CONFIG.DIAS_ALERTA * 24 * 3600 * 1000);

  const linhas = contas.getRange(2, 1, contas.getLastRow() - 1, CAB_CONTAS.length).getValues()
    .filter(r => r[5] !== 'Pago')
    .filter(r => !(r[2] instanceof Date) || r[2] <= limite);
  if (!linhas.length) return;

  const html = linhas.map(r => {
    const venc = r[2] instanceof Date ? r[2] : null;
    const atrasado = venc && venc < hoje;
    const cor = atrasado ? '#c0392b' : '#333';
    return '<tr style="color:' + cor + '"><td>' + r[0] + '</td><td>' +
      (venc ? Utilities.formatDate(venc, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '?') +
      (atrasado ? ' (VENCIDO)' : '') + '</td><td>' + (r[3] ? moeda(r[3]) : '?') +
      '</td><td style="font-family:monospace;font-size:12px">' + (r[4] || '') +
      '</td><td><a href="' + r[7] + '">abrir</a></td></tr>';
  }).join('');

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: '💰 ' + linhas.length + ' boleto(s) vencendo ou vencido(s)',
    htmlBody: '<p>Contas a pagar nos próximos ' + CONFIG.DIAS_ALERTA + ' dias:</p>' +
      '<table cellpadding="6" style="border-collapse:collapse" border="1">' +
      '<tr><th>Fornecedor</th><th>Vencimento</th><th>Valor</th><th>Linha digitável</th><th>E-mail</th></tr>' +
      html + '</table><p><a href="' + (ScriptApp.getService().getUrl() || planilha().getUrl()) +
      '">Abrir Meus Boletos</a> — marque como "Pago" depois de pagar.</p>'
  });
}

function moeda(v) {
  return 'R$ ' + Number(v).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
