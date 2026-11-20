/**
 * Cadastro de Visitantes · Batizado Projeto Brincapoeira (Capoeira Porto Bahia)
 * Recebe cada cadastro do site de visitantes e grava numa planilha Google SEPARADA
 * (exportável para Excel: Arquivo → Fazer download → Microsoft Excel).
 *
 * Aba criada automaticamente: "Visitantes" → 1 linha por visitante,
 * com coluna de presença (caixinha) para marcar na portaria.
 *
 * Como instalar: veja o arquivo LEIAME.
 */

var ABA = 'Visitantes';
var FUSO = 'America/Sao_Paulo';
var CABECALHO = ['Data/hora', 'Protocolo', 'Nome', 'Telefone', 'CPF', 'Grupo', 'Presença'];
var COR = '#2F3BC4';
var LARGURAS = [150, 115, 250, 145, 135, 230, 95];
// Planilha "visitantes brincapoeira" — usada se o script não estiver dentro dela
var PLANILHA_ID = '1jO9aUHQwGGKkoLPbHRbgs-9lG3zjBgY5QohQgKmF9qo';

function planilha() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PLANILHA_ID);
}

/* ------------------------------------------------------------------ */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return resposta({ ok: false, erro: 'Muitos envios ao mesmo tempo. Tente de novo em alguns segundos.' });
  }
  try {
    var dados = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var problema = validar(dados);
    if (problema) return resposta({ ok: false, erro: problema });

    var carimbo = Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy HH:mm:ss');
    var linha = montarLinha(dados, carimbo);

    var aba = garantirAba(planilha());
    if (protocoloJaExiste(aba, linha.protocolo)) {
      return resposta({ ok: true, protocolo: linha.protocolo, duplicado: true });
    }

    var ini = aba.getLastRow() + 1;
    if (ini > aba.getMaxRows()) aba.insertRowsAfter(aba.getMaxRows(), 200);
    aba.getRange(ini, 1, 1, linha.valores.length).setValues([linha.valores]);
    aba.getRange(ini, CABECALHO.length).insertCheckboxes();

    return resposta({ ok: true, protocolo: linha.protocolo });
  } catch (err) {
    return resposta({ ok: false, erro: 'Erro ao salvar: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

/** Abrir a URL /exec no navegador mostra só isto (não expõe nenhum dado). */
function doGet() {
  return ContentService.createTextOutput('Cadastro de Visitantes Brincapoeira: online ✔');
}

/* ------------------------- regras puras ---------------------------- */

function validar(d) {
  if (!d || typeof d !== 'object') return 'Dados inválidos.';
  if (limpo(d.nome).split(' ').length < 2) return 'Informe nome e sobrenome.';
  if (limpo(d.telefone).replace(/\D/g, '').length < 10) return 'Informe um telefone com DDD.';
  if (!cpfValido(d.cpf)) return 'CPF inválido.';
  if (!limpo(d.grupo)) return 'Informe o grupo de capoeira.';
  return '';
}

function cpfValido(v) {
  var d = limpo(v).replace(/\D/g, '');
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  for (var t = 9; t < 11; t++) {
    var s = 0;
    for (var i = 0; i < t; i++) s += +d[i] * (t + 1 - i);
    if (((s * 10) % 11) % 10 !== +d[t]) return false;
  }
  return true;
}

/** Função pura (sem acesso à planilha) — dá para testar fora do Google. */
function montarLinha(d, carimbo) {
  var protocolo = limpo(d.protocolo) || ('VS-' + Math.random().toString(36).slice(2, 8).toUpperCase());
  var cpf = limpo(d.cpf).replace(/\D/g, '').replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
  var valores = [carimbo, protocolo, limpo(d.nome), limpo(d.telefone), cpf, limpo(d.grupo)];
  return { protocolo: protocolo, valores: valores.map(seguro) };
}

function limpo(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, 200);
}

/** Força texto: evita fórmulas ("=...") e que números percam zeros/formatação. */
function seguro(v) {
  if (typeof v !== 'string' || v === '') return v;
  return /^[=+\-@']/.test(v) || /^[\d\s.,:\/()\-]+$/.test(v) ? "'" + v : v;
}

/* ------------------------- planilha -------------------------------- */

function garantirAba(ss) {
  var aba = ss.getSheetByName(ABA);
  var nova = !aba;
  if (nova) {
    // aproveita a primeira aba vazia ("Página1") se existir
    var primeira = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && primeira.getLastRow() === 0) { aba = primeira; aba.setName(ABA); }
    else aba = ss.insertSheet(ABA);
  }
  if (aba.getMaxColumns() < CABECALHO.length) {
    aba.insertColumnsAfter(aba.getMaxColumns(), CABECALHO.length - aba.getMaxColumns());
  }
  var faixa = aba.getRange(1, 1, 1, CABECALHO.length);
  if (nova || faixa.getDisplayValues()[0].join('|') !== CABECALHO.join('|')) {
    faixa.setValues([CABECALHO]);
    formatar(aba);
  }
  return aba;
}

/** Visual da planilha: cabeçalho azul, linhas zebradas, filtro e verde em quem já entrou. */
function formatar(aba) {
  var n = CABECALHO.length;
  if (aba.getMaxColumns() > n) aba.deleteColumns(n + 1, aba.getMaxColumns() - n);
  if (aba.getMaxRows() < 300) aba.insertRowsAfter(aba.getMaxRows(), 300 - aba.getMaxRows());
  var linhas = aba.getMaxRows();
  var tudo = aba.getRange(1, 1, linhas, n);

  aba.setHiddenGridlines(true);
  tudo.setFontFamily('Nunito').setFontSize(11).setFontColor('#1D2A22')
    .setVerticalAlignment('middle').setWrap(false);
  LARGURAS.forEach(function (w, i) { aba.setColumnWidth(i + 1, w); });
  // datas, protocolo, telefone, CPF e presença centralizados
  [1, 2, 4, 5, 7].forEach(function (c) { aba.getRange(1, c, linhas, 1).setHorizontalAlignment('center'); });
  aba.getRange(2, 3, linhas - 1, 1).setFontWeight('bold');
  aba.getRange(2, 2, linhas - 1, 1).setFontColor('#E3262B').setFontWeight('bold');
  aba.getRange(2, 1, linhas - 1, 1).setFontColor('#5C6A61').setFontSize(10);

  aba.getBandings().forEach(function (b) { b.remove(); });
  tudo.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false)
    .setHeaderRowColor(COR).setFirstRowColor('#FFFFFF').setSecondRowColor('#F1F3FD');

  aba.getRange(1, 1, 1, n).setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11)
    .setHorizontalAlignment('center');
  aba.setFrozenRows(1);
  aba.setRowHeight(1, 38);
  aba.setRowHeights(2, linhas - 1, 28);
  tudo.setBorder(null, null, true, null, null, true, '#E1E4F5', SpreadsheetApp.BorderStyle.SOLID);

  // quem já entrou (caixinha marcada) fica verde
  var dados = aba.getRange(2, 1, linhas - 1, n);
  aba.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2=TRUE')
      .setBackground('#DFF3E4').setFontColor('#0E5A33')
      .setRanges([dados]).build()
  ]);

  if (aba.getFilter()) aba.getFilter().remove();
  tudo.createFilter();
}

function protocoloJaExiste(aba, protocolo) {
  var n = aba.getLastRow() - 1;
  if (n < 1) return false;
  return aba.getRange(2, 2, n, 1).getDisplayValues().some(function (l) { return l[0] === protocolo; });
}

/** Menu "Brincapoeira" na planilha. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Brincapoeira')
    .addItem('Arrumar e deixar bonita', 'arrumarPlanilha')
    .addToUi();
}

function arrumarPlanilha() {
  var ss = planilha();
  var aba = garantirAba(ss);
  formatar(aba);
  var removidas = 0, ultima = aba.getLastRow();
  if (ultima >= 2) {
    var nomes = aba.getRange(2, 3, ultima - 1, 1).getDisplayValues();
    for (var i = nomes.length - 1; i >= 0; i--) {
      if (!String(nomes[i][0]).trim()) { aba.deleteRow(i + 2); removidas++; }
    }
  }
  try { ss.toast('Cabeçalho recolocado e ' + removidas + ' linha(s) vazia(s) removida(s).', 'Brincapoeira', 6); } catch (e) {}
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------- teste manual ---------------------------- */

/** Rode no editor (botão ▶) para ver se a planilha recebe uma linha. Depois apague a linha TESTE. */
function testeManual() {
  var r = doPost({ postData: { contents: JSON.stringify({
    protocolo: 'VS-TESTE' + Math.floor(Math.random() * 1000),
    nome: 'Visitante TESTE', telefone: '(51) 99999-9999', cpf: '529.982.247-25', grupo: 'Grupo Teste'
  }) } });
  Logger.log(r.getContent());
}
