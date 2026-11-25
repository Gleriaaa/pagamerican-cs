/* ============================================================
   Reenvio · planilhas das fábricas (EagleLabs, ShipOffers, RedRock)
   ------------------------------------------------------------
   Web App separado do script da aba "Reenvios" (reenvio.gs).
   - GET  ?chave=...&order=183156  → tracking, status e data do reenvio
                                     (nunca devolve e-mail nem endereço)
   - POST {chave, acao:"adicionar", ...} → com a validação ligada, o
                                     reenvio vai para a planilha interna
                                     "Validação de reenvios" e a validadora
                                     recebe um e-mail. Só depois de
                                     "Aprovado" ele entra na planilha da
                                     fábrica e a fábrica é avisada.
                                     Sem validação: vai direto para a fábrica.
   - aoEditarValidacao / varrerValidacao → processam a coluna Decisão
                                     (configurarValidacao instala os dois)
   - lembreteDiario() → todo dia, e-mail para cada fábrica com os
                        reenvios ainda sem tracking (instalarLembrete)
   Como publicar: ver apps-script/README-reenvio-fabricas.md
   ============================================================ */

const CHAVE = "pa-trk-sbkaevI2NaUoIYQ1nsHi";

/* Uma entrada por fábrica: o ID é o trecho do link da planilha entre /d/ e /edit.
   "aba" vazio = primeira aba da planilha. */
const FABRICAS = {
  eaglelabs:  { nome: "EagleLabs",  id: "1pLsXlK9vMCcrKqcjpCoRxzYb0Ql-MAJOzKn6lrwdxRY",  aba: "" },
  shipoffers: { nome: "ShipOffers", id: "1j5IWjp2yUqwox8tnfiyl4wGYLydSC-LowAvqqg9Fcp8", aba: "" },
  redrock:    { nome: "RedRock",    id: "1PLXauG7pAx_Qszf0RnmfWeFI3g9Tw7qLS983DWb0QVc",    aba: "" }
};

const STATUS_INICIAL = "Sent to factory";

/* Avisos por e-mail (em inglês). Quem recebe: todos os EDITORES de cada planilha, menos o dono.
   Ou seja, compartilhou a planilha com a fábrica como Editor → ela passa a receber os avisos.
   O e-mail nunca traz endereço nem e-mail do cliente: só pedido, produto e o link da planilha. */
const AVISO_COPIA = [];            /* e-mails da PagAmerican que recebem cópia, ex.: ["gabriel.gleria@pagamerican.com.br"] */
const LEMBRETE_DIAS_MIN = 2;       /* lembrete diário lista reenvios sem tracking há pelo menos N dias */
const LEMBRETE_HORA = 9;           /* hora do lembrete, no fuso abaixo */
const LEMBRETE_FUSO = "America/New_York";
const ASSINATURA = "PagAmerican Support";

/* Validação antes da fábrica (planilha interna, as fábricas NÃO têm acesso).
   A planilha é criada pela função configurarValidacao e o ID fica guardado no projeto;
   só preencha VALIDACAO_ID se quiser usar uma planilha que já existe.
   Quem recebe o aviso de "reenvio para validar": todos os EDITORES dessa planilha, menos o dono. */
const VALIDACAO_ID = "";
const VALIDACAO_COPIA = [];        /* e-mails que recebem cópia do aviso de validação */
const VALIDACAO_VARREDURA_MIN = 10; /* de quantos em quantos minutos a varredura confere as decisões */
const VALIDACAO_FUSO = "America/Sao_Paulo";

/* Cabeçalhos procurados, em português ou inglês (sem acento, minúsculo, só letras e números). */
const COLS = {
  order:      ["ordernumber", "order", "pedido"],
  produto:    ["produto", "product"],
  quantidade: ["quantidade", "qtd", "quantity"],
  endereco:   ["enderecodeenvio", "shippingaddress", "endereco", "address"],
  email:      ["customeremail", "email"],
  nome:       ["customername", "nomedocliente", "nome", "name"],
  vendor:     ["accountvendor", "vendor", "account", "conta"],
  protocolo:  ["protocolo", "protocol"],
  ticket:     ["ticket"],
  dataEnvio:  ["datadoenvioafabrica", "datesenttofactory", "dataenvio", "datesent", "data", "date"],
  tracking:   ["trackingdoreenvio", "reshipmenttracking", "trackingnumber", "tracking", "rastreio"],
  status:     ["status"],
  obs:        ["observacoes", "notes", "obs", "note"]
};

function norm_(s) {
  return String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[^a-z0-9]/g, "");
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function orderKey_(s) { return String(s == null ? "" : s).replace(/[^0-9a-z-]/gi, "").toLowerCase(); }

function aba_(f) {
  let ss;
  try { ss = SpreadsheetApp.openById(f.id); }
  catch (err) {
    throw new Error("Não consegui abrir a planilha da " + f.nome + ". Se ela ainda for um arquivo Excel (.xlsx), abra e use Arquivo → Salvar como Planilhas Google, e coloque o ID novo no script.");
  }
  return f.aba ? ss.getSheetByName(f.aba) : ss.getSheets()[0];
}

/* Acha a linha de cabeçalho (até a 5ª) e o índice de cada coluna. */
function mapa_(sh) {
  const top = sh.getRange(1, 1, Math.min(5, sh.getLastRow() || 1), sh.getLastColumn() || 1).getDisplayValues();
  for (let r = 0; r < top.length; r++) {
    const heads = top[r].map(norm_);
    if (!heads.some(h => COLS.order.indexOf(h) >= 0)) continue;
    const idx = {};
    Object.keys(COLS).forEach(k => {
      /* exato primeiro; depois "começa com" (ex.: "trackingdoreenvio") */
      let i = heads.findIndex(h => COLS[k].indexOf(h) >= 0);
      if (i < 0) i = heads.findIndex(h => h && COLS[k].some(c => h.indexOf(c) === 0));
      idx[k] = i;
    });
    return { linha: r + 1, idx: idx, ncols: heads.length };
  }
  throw new Error("Cabeçalho com 'Order Number' não encontrado na planilha.");
}

function fabrica_(nome) {
  const n = norm_(nome);
  return Object.keys(FABRICAS).map(k => FABRICAS[k]).find(f => norm_(f.nome) === n) || null;
}

/* ---------- consulta ---------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.chave !== CHAVE) return json_({ ok: false, erro: "Chave inválida." });
    const alvo = orderKey_(p.order);
    if (!alvo) return json_({ ok: false, erro: "Informe o número do pedido." });

    const resultados = [];
    Object.keys(FABRICAS).forEach(k => {
      const f = FABRICAS[k];
      if (!f.id || f.id.indexOf("COLE_") === 0) return;
      const sh = aba_(f), m = mapa_(sh), last = sh.getLastRow();
      if (last <= m.linha) return;
      const vals = sh.getRange(m.linha + 1, 1, last - m.linha, m.ncols).getDisplayValues();
      const g = (row, c) => (m.idx[c] >= 0 ? row[m.idx[c]] : "");
      vals.forEach(row => {
        if (orderKey_(g(row, "order")) !== alvo) return;
        resultados.push({
          fabrica: f.nome,
          protocolo: g(row, "protocolo"),
          ticket: g(row, "ticket"),
          dataEnvio: g(row, "dataEnvio"),
          produto: g(row, "produto"),
          quantidade: g(row, "quantidade"),
          tracking: g(row, "tracking"),
          status: g(row, "status"),
          observacoes: g(row, "obs")
        });
      });
    });
    /* reenvios ainda na validação (ou reprovados): os aprovados já aparecem pela fábrica */
    try { buscarNaValidacao_(alvo).forEach(r => resultados.push(r)); }
    catch (err) { console.error("Consulta à validação falhou: " + err); }
    return json_({ ok: true, order: p.order, resultados: resultados });
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  }
}

/* ---------- novo reenvio na planilha da fábrica ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (d.chave !== CHAVE) return json_({ ok: false, erro: "Chave inválida." });
    if (d.acao !== "adicionar") return json_({ ok: false, erro: "Ação desconhecida." });
    lock.waitLock(20000);
    if (idValidacao_()) return json_(solicitarValidacao_(d));
    return json_(gravarNaFabrica_(d));
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* Grava o reenvio na planilha da fábrica e avisa a fábrica. Quem chama segura o lock. */
function gravarNaFabrica_(d) {
  const f = fabrica_(d.fabrica);
  if (!f || !f.id || f.id.indexOf("COLE_") === 0)
    return { ok: false, semPlanilha: true, erro: "A fábrica " + (d.fabrica || "?") + " ainda não tem planilha de acompanhamento." };
  {
    const sh = aba_(f), m = mapa_(sh), last = sh.getLastRow();

    /* não duplica o mesmo protocolo */
    if (d.protocolo && m.idx.protocolo >= 0 && last > m.linha) {
      const protos = sh.getRange(m.linha + 1, m.idx.protocolo + 1, last - m.linha, 1).getDisplayValues().map(r => r[0]);
      if (protos.indexOf(d.protocolo) >= 0) return { ok: true, fabrica: f.nome, jaExistia: true };
    }

    const row = new Array(m.ncols).fill("");
    const put = (c, v) => { if (m.idx[c] >= 0) row[m.idx[c]] = v == null ? "" : String(v); };
    /* apóstrofo = gravar como texto, igual às linhas que já existem (senão vira número/data) */
    put("order", d.orderNumber ? "'" + d.orderNumber : "");
    put("produto", d.produto);
    put("quantidade", d.quantidade);
    put("endereco", d.endereco);
    put("email", d.email);
    put("nome", d.nomeCliente);
    put("vendor", d.vendor);
    put("protocolo", d.protocolo);
    put("ticket", d.ticket ? "'" + d.ticket : "");
    put("dataEnvio", "'" + Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy"));
    put("status", STATUS_INICIAL);
    sh.appendRow(row);
    /* a linha nova herda cores e bordas da linha de cima */
    const nova = sh.getLastRow();
    if (nova > m.linha + 1)
      sh.getRange(nova - 1, 1, 1, m.ncols).copyTo(sh.getRange(nova, 1, 1, m.ncols), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    let avisados = 0;
    try { avisados = avisarNovo_(f, sh, d); } catch (err) { console.error("Aviso por e-mail falhou: " + err); }
    return { ok: true, fabrica: f.nome, avisados: avisados };
  }
}

/* ---------- avisos por e-mail ---------- */
function destinatarios_(sh) {
  const dono = Session.getEffectiveUser().getEmail().toLowerCase();
  return sh.getParent().getEditors().map(u => u.getEmail().toLowerCase())
    .filter(e => e && e !== dono);
}
function linkPlanilha_(f) { return "https://docs.google.com/spreadsheets/d/" + f.id + "/edit"; }
function esc_(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

function enviar_(para, assunto, texto, html, copia, nome) {
  if (!para.length) return 0;
  const cc = copia || AVISO_COPIA;
  const opt = { htmlBody: html, name: nome || ASSINATURA };
  if (cc.length) opt.cc = cc.join(",");
  MailApp.sendEmail(para.join(","), assunto, texto, opt);
  return para.length;
}

/* Novo reenvio → e-mail na hora para a fábrica. */
function avisarNovo_(f, sh, d) {
  const para = destinatarios_(sh);
  const link = linkPlanilha_(f);
  const assunto = "New reshipment request – Order " + d.orderNumber + " (" + f.nome + ")";
  const linhas = [["Order Number", d.orderNumber], ["Product", d.produto], ["Quantity", d.quantidade], ["Protocol", d.protocolo]];
  const texto = "Hi " + f.nome + " team,\n\nA new reshipment request was added to your sheet:\n\n" +
    linhas.map(l => l[0] + ": " + (l[1] || "-")).join("\n") +
    "\n\nThe shipping address is in the sheet. Once it ships, please fill in the Reshipment tracking and the Status:\n" + link +
    "\n\nThank you!\n" + ASSINATURA;
  const html = "<p>Hi " + esc_(f.nome) + " team,</p><p>A new reshipment request was added to your sheet:</p>" +
    "<table cellpadding='4' style='border-collapse:collapse'>" +
    linhas.map(l => "<tr><td style='color:#666'>" + l[0] + "</td><td><b>" + esc_(l[1] || "-") + "</b></td></tr>").join("") +
    "</table><p>The shipping address is in the sheet. Once it ships, please fill in the <b>Reshipment tracking</b> and the <b>Status</b>.</p>" +
    "<p><a href='" + link + "'>Open the " + esc_(f.nome) + " reshipment sheet</a></p><p>Thank you!<br>" + ASSINATURA + "</p>";
  return enviar_(para, assunto, texto, html);
}

/* Datas da planilha vêm como "dd/MM/yyyy". */
function lerData_(txt) {
  const m = String(txt || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}
function diasDesde_(txt) {
  const dt = lerData_(txt);
  return dt ? Math.floor((Date.now() - dt.getTime()) / 86400000) : null;
}
/* "02/10/2026" → "Oct 2, 2026", para a fábrica (EUA) não ler como 10 de fevereiro. */
function dataEn_(txt) {
  const dt = lerData_(txt);
  return dt ? ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][dt.getMonth()] + " " + dt.getDate() + ", " + dt.getFullYear() : txt;
}

/* Todo dia: para cada fábrica, os reenvios sem tracking (e não entregues/cancelados) há LEMBRETE_DIAS_MIN dias ou mais. */
function lembreteDiario() {
  Object.keys(FABRICAS).forEach(k => {
    const f = FABRICAS[k];
    if (!f.id || f.id.indexOf("COLE_") === 0) return;
    try {
      const sh = aba_(f), m = mapa_(sh), last = sh.getLastRow();
      const para = destinatarios_(sh);
      if (!para.length || last <= m.linha) return;
      const vals = sh.getRange(m.linha + 1, 1, last - m.linha, m.ncols).getDisplayValues();
      const g = (row, c) => (m.idx[c] >= 0 ? row[m.idx[c]] : "");
      const pend = vals.map(row => ({
        order: g(row, "order"), produto: g(row, "produto"), qtd: g(row, "quantidade"),
        data: dataEn_(g(row, "dataEnvio")), dias: diasDesde_(g(row, "dataEnvio")), status: g(row, "status"),
        tracking: g(row, "tracking")
      })).filter(r => r.order && !r.tracking &&
        !/deliver|entregue|cancel/.test(norm_(r.status)) &&
        (r.dias == null || r.dias >= LEMBRETE_DIAS_MIN))
        .sort((a, b) => (b.dias || 0) - (a.dias || 0));
      if (!pend.length) return;

      const MAX = 50, lista = pend.slice(0, MAX), resto = pend.length - lista.length;
      const link = linkPlanilha_(f);
      const assunto = "Reminder: " + pend.length + " reshipment" + (pend.length > 1 ? "s" : "") + " waiting for tracking (" + f.nome + ")";
      const texto = "Hi " + f.nome + " team,\n\nThese reshipments still have no tracking number:\n\n" +
        lista.map(r => "- Order " + r.order + " | " + r.produto + " (" + r.qtd + ") | sent " + r.data + (r.dias != null ? " (" + r.dias + " days ago)" : "") + " | " + (r.status || "no status")).join("\n") +
        (resto > 0 ? "\n...and " + resto + " more." : "") +
        "\n\nPlease fill in the Reshipment tracking and the Status:\n" + link + "\n\nThank you!\n" + ASSINATURA;
      const html = "<p>Hi " + esc_(f.nome) + " team,</p><p>These reshipments still have no tracking number:</p>" +
        "<table cellpadding='5' style='border-collapse:collapse;font-size:13px'><tr style='background:#f2f4f7'><th align='left'>Order</th><th align='left'>Product</th><th align='left'>Qty</th><th align='left'>Sent to factory</th><th align='left'>Status</th></tr>" +
        lista.map(r => "<tr style='border-top:1px solid #e4e7ec'><td><b>" + esc_(r.order) + "</b></td><td>" + esc_(r.produto) + "</td><td>" + esc_(r.qtd) + "</td><td>" + esc_(r.data) + (r.dias != null ? " <span style='color:#b54708'>(" + r.dias + " days)</span>" : "") + "</td><td>" + esc_(r.status || "-") + "</td></tr>").join("") +
        "</table>" + (resto > 0 ? "<p>...and " + resto + " more in the sheet.</p>" : "") +
        "<p><a href='" + link + "'>Open the " + esc_(f.nome) + " reshipment sheet</a> and fill in the <b>Reshipment tracking</b> and the <b>Status</b>.</p><p>Thank you!<br>" + ASSINATURA + "</p>";
      enviar_(para, assunto, texto, html);
    } catch (err) {
      console.error(f.nome + ": " + err);
    }
  });
}

/* Rode UMA vez pelo editor (selecione instalarLembrete e clique em Executar) para ligar o lembrete diário. */
function instalarLembrete() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "lembreteDiario").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("lembreteDiario").timeBased().everyDays(1).atHour(LEMBRETE_HORA).inTimezone(LEMBRETE_FUSO).create();
}

/* Rode pelo editor para conferir os avisos: libera a permissão de enviar e-mail (na 1ª vez),
   manda para VOCÊ um exemplo do aviso e mostra no registro quem recebe em cada fábrica. */
function testarAviso() {
  const eu = Session.getEffectiveUser().getEmail();
  Object.keys(FABRICAS).forEach(k => {
    const f = FABRICAS[k];
    try { console.log(f.nome + " → recebem: " + (destinatarios_(aba_(f)).join(", ") || "ninguém (compartilhe a planilha como Editor)")); }
    catch (err) { console.log(f.nome + " → erro: " + err); }
  });
  const f = FABRICAS.redrock;
  const exemplo = { orderNumber: "TEST-0001", produto: "Sample product", quantidade: "1", protocolo: "REE-TEST" };
  MailApp.sendEmail(eu, "[TEST] New reshipment request – Order TEST-0001 (" + f.nome + ")",
    "Teste do aviso de reenvio. Se este e-mail chegou, o envio de avisos está funcionando.",
    { name: ASSINATURA, htmlBody: "<p><b>Teste do aviso de reenvio.</b> Se este e-mail chegou, o envio de avisos está funcionando.</p><hr>" +
      "<p>Order Number: <b>" + exemplo.orderNumber + "</b><br>Product: <b>" + exemplo.produto + "</b><br>Protocol: <b>" + exemplo.protocolo + "</b></p>" +
      "<p><a href='" + linkPlanilha_(f) + "'>Open the " + f.nome + " reshipment sheet</a></p>" });
  console.log("E-mail de teste enviado para " + eu + ". Cotas restantes hoje: " + MailApp.getRemainingDailyQuota());
}

/* ============================================================
   Validação antes da fábrica
   ------------------------------------------------------------
   Planilha interna "Validação de reenvios" (só o time PagAmerican).
   - O site manda o reenvio para cá, com "Aguardando validação".
   - A validadora recebe um e-mail e escolhe Aprovado ou Reprovado
     na coluna Decisão. Só as colunas Decisão e Motivo da reprovação
     ficam liberadas para edição; o resto é protegido.
   - Aprovado → o script grava na planilha da fábrica e avisa a fábrica.
   - Reprovado → exige o motivo; nada vai para a fábrica.
   ============================================================ */

const VCOLS = [
  ["protocolo", "Protocolo", 150], ["recebido", "Recebido em", 130], ["agente", "Agente", 120],
  ["origem", "Origem", 90], ["ticket", "Ticket", 80], ["order", "Order Number", 110],
  ["fabrica", "Fábrica", 100], ["produto", "Produto", 220], ["quantidade", "Quantidade", 90],
  ["motivo", "Motivo", 190], ["obs", "Observações do agente", 240], ["anteriores", "Reenvios anteriores", 190],
  ["decisao", "Decisão", 120], ["motivoReprov", "Motivo da reprovação", 240], ["situacao", "Situação", 260],
  ["validadoPor", "Validado por", 200], ["validadoEm", "Validado em", 130],
  ["nome", "Nome do cliente", 170], ["endereco", "Endereço de envio", 320], ["email", "E-mail do cliente", 220],
  ["vendor", "Vendor", 160]
];
const V_ABA = "Validação";
const V_APROVADO = "Aprovado", V_REPROVADO = "Reprovado", V_AGUARDANDO = "Aguardando validação";
const V_COR = { espera: "#fef7e0", ok: "#e6f4ea", ruim: "#fce8e6", alerta: "#fff4e5", nada: null };

function idValidacao_() {
  return VALIDACAO_ID || PropertiesService.getScriptProperties().getProperty("VALIDACAO_ID") || "";
}
function abaValidacao_() {
  const ss = SpreadsheetApp.openById(idValidacao_());
  return ss.getSheetByName(V_ABA) || ss.getSheets()[0];
}
/* índice (0-based) de cada coluna pelo nome do cabeçalho na linha 1 */
function vmapa_(sh) {
  const heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(norm_);
  const idx = {};
  VCOLS.forEach(c => { idx[c[0]] = heads.indexOf(norm_(c[1])); });
  ["order", "fabrica", "decisao", "situacao"].forEach(k => {
    if (idx[k] < 0) throw new Error("Coluna \"" + VCOLS.find(c => c[0] === k)[1] + "\" não encontrada na planilha de validação. Não renomeie as colunas.");
  });
  return { idx: idx, ncols: heads.length };
}
function agora_() { return Utilities.formatDate(new Date(), VALIDACAO_FUSO, "dd/MM/yyyy HH:mm"); }
function linkValidacao_(sh, linha) {
  return "https://docs.google.com/spreadsheets/d/" + idValidacao_() + "/edit#gid=" + sh.getSheetId() + (linha ? "&range=A" + linha : "");
}

/* Quantos reenvios essa order já tem (planilhas das fábricas + validação ainda não enviada). */
function contarAnteriores_(orderNum, protocolo) {
  const alvo = orderKey_(orderNum), partes = [];
  if (!alvo) return "";
  Object.keys(FABRICAS).forEach(k => {
    const f = FABRICAS[k];
    if (!f.id || f.id.indexOf("COLE_") === 0) return;
    try {
      const sh = aba_(f), m = mapa_(sh), last = sh.getLastRow();
      if (last <= m.linha || m.idx.order < 0) return;
      const n = sh.getRange(m.linha + 1, m.idx.order + 1, last - m.linha, 1).getDisplayValues()
        .filter(r => orderKey_(r[0]) === alvo).length;
      if (n) partes.push(n + " na " + f.nome);
    } catch (err) { console.error("Contagem na " + f.nome + ": " + err); }
  });
  try {
    const sh = abaValidacao_(), m = vmapa_(sh), last = sh.getLastRow();
    if (last > 1) {
      const vals = sh.getRange(2, 1, last - 1, m.ncols).getDisplayValues();
      const n = vals.filter(r => orderKey_(r[m.idx.order]) === alvo &&
        (m.idx.protocolo < 0 || r[m.idx.protocolo] !== protocolo) &&
        !/^Enviado/.test(r[m.idx.situacao])).length;   /* os enviados já contam pela fábrica */
      if (n) partes.push(n + " na validação");
    }
  } catch (err) { console.error("Contagem na validação: " + err); }
  return partes.length ? "⚠ " + partes.join(", ") : "Nenhum";
}

/* Site → nova linha na validação + e-mail para a validadora. Quem chama segura o lock. */
function solicitarValidacao_(d) {
  const sh = abaValidacao_(), m = vmapa_(sh), last = sh.getLastRow();
  if (d.protocolo && m.idx.protocolo >= 0 && last > 1) {
    const protos = sh.getRange(2, m.idx.protocolo + 1, last - 1, 1).getDisplayValues().map(r => r[0]);
    if (protos.indexOf(d.protocolo) >= 0) return { ok: true, validacao: true, jaExistia: true, fabrica: d.fabrica };
  }
  const anteriores = contarAnteriores_(d.orderNumber, d.protocolo);
  const row = new Array(m.ncols).fill("");
  const put = (c, v) => { if (m.idx[c] >= 0) row[m.idx[c]] = v == null ? "" : String(v); };
  const texto = v => (v ? "'" + v : "");   /* apóstrofo = texto (order e ticket não viram número) */
  put("protocolo", d.protocolo); put("recebido", texto(agora_())); put("agente", d.agente);
  put("origem", d.origem); put("ticket", texto(d.ticket)); put("order", texto(d.orderNumber));
  put("fabrica", d.fabrica); put("produto", d.produto); put("quantidade", texto(d.quantidade));
  put("motivo", d.motivo); put("obs", d.observacoes); put("anteriores", anteriores);
  put("situacao", V_AGUARDANDO); put("nome", d.nomeCliente); put("endereco", d.endereco);
  put("email", d.email); put("vendor", d.vendor);
  sh.appendRow(row);
  const linha = sh.getLastRow();
  sh.getRange(linha, m.idx.situacao + 1).setBackground(V_COR.espera);
  if (m.idx.anteriores >= 0) sh.getRange(linha, m.idx.anteriores + 1).setBackground(/^⚠/.test(anteriores) ? V_COR.alerta : null);
  let avisados = 0;
  try { avisados = avisarValidacao_(sh, linha, d, anteriores); } catch (err) { console.error("Aviso de validação falhou: " + err); }
  return { ok: true, validacao: true, fabrica: d.fabrica, avisados: avisados, anteriores: anteriores };
}

/* E-mail para quem valida (editores da planilha de validação, menos o dono). Sem endereço nem e-mail do cliente. */
function avisarValidacao_(sh, linha, d, anteriores) {
  let para = destinatarios_(sh), copia = VALIDACAO_COPIA;
  if (!para.length && copia.length) { para = copia; copia = []; }
  if (!para.length) { console.warn("Ninguém recebe o aviso de validação: compartilhe a planilha como Editor."); return 0; }
  const link = linkValidacao_(sh, linha);
  const assunto = "Reenvio para validar – Order " + d.orderNumber + " (" + (d.fabrica || "?") + ")";
  const linhas = [["Protocolo", d.protocolo], ["Agente", d.agente], ["Origem", [d.origem, d.ticket && "ticket #" + String(d.ticket).replace(/^#/, "")].filter(Boolean).join(" · ")],
    ["Order Number", d.orderNumber], ["Fábrica", d.fabrica], ["Produto", d.produto], ["Quantidade", d.quantidade],
    ["Motivo", d.motivo], ["Observações", d.observacoes], ["Reenvios anteriores", anteriores]];
  const texto = "Olá!\n\nChegou um reenvio para validar:\n\n" +
    linhas.map(l => l[0] + ": " + (l[1] || "-")).join("\n") +
    "\n\nAbra a planilha e escolha Aprovado ou Reprovado na coluna Decisão. Ele só vai para a fábrica depois de aprovado.\n" + link;
  const html = "<p>Olá!</p><p>Chegou um reenvio para validar:</p>" +
    "<table cellpadding='4' style='border-collapse:collapse'>" +
    linhas.map(l => "<tr><td style='color:#666'>" + l[0] + "</td><td><b" + (l[0] === "Reenvios anteriores" && /^⚠/.test(l[1]) ? " style='color:#b54708'" : "") + ">" + esc_(l[1] || "-") + "</b></td></tr>").join("") +
    "</table><p>Abra a planilha e escolha <b>Aprovado</b> ou <b>Reprovado</b> na coluna <b>Decisão</b>. Ele só vai para a fábrica depois de aprovado.</p>" +
    "<p><a href='" + link + "'>Abrir a linha na planilha de validação</a></p>";
  return enviar_(para, assunto, texto, html, copia, "PagAmerican · Reenvios");
}

/* Aplica a decisão de uma linha. Quem chama segura o lock. */
function processarLinha_(sh, m, linha, quem) {
  const row = sh.getRange(linha, 1, 1, m.ncols).getDisplayValues()[0];
  const g = k => (m.idx[k] >= 0 ? row[m.idx[k]] : "");
  const cel = k => sh.getRange(linha, m.idx[k] + 1);
  const set = (k, v) => { if (m.idx[k] >= 0) cel(k).setValue(v); };
  const situacao = (txt, cor) => { cel("situacao").setValue(txt).setBackground(cor); };
  if (!g("order") && !g("protocolo")) return;
  const dec = norm_(g("decisao")), sit = g("situacao");

  /* já foi para a fábrica: a decisão não muda mais por aqui */
  if (/^Enviado/.test(sit)) {
    if (dec !== norm_(V_APROVADO)) {
      cel("decisao").setValue(V_APROVADO);
      cel("decisao").setNote("Este reenvio já foi enviado à fábrica em " + (g("validadoEm") || "?") +
        ". Para cancelar, fale com a fábrica e registre em Observações.");
    }
    return;
  }

  if (dec === norm_(V_APROVADO)) {
    const d = {
      protocolo: g("protocolo"), orderNumber: g("order"), produto: g("produto"), quantidade: g("quantidade"),
      endereco: g("endereco"), email: g("email"), nomeCliente: g("nome"), vendor: g("vendor"),
      ticket: g("ticket"), fabrica: g("fabrica")
    };
    let r;
    try { r = gravarNaFabrica_(d); } catch (err) { r = { ok: false, erro: String(err && err.message || err) }; }
    if (r.ok) situacao("Enviado à " + r.fabrica + (r.jaExistia ? " (já estava lá)" : "") + " em " + agora_(), V_COR.ok);
    else if (r.semPlanilha) situacao("Aprovado: a " + (d.fabrica || "fábrica") + " não tem planilha. Enviar manualmente.", V_COR.alerta);
    else { situacao("Erro ao enviar: " + r.erro + " (tenta de novo sozinho)", V_COR.ruim); return; }
    if (!g("validadoPor")) set("validadoPor", quem || "—");
    set("validadoEm", "'" + agora_());
    return;
  }

  if (dec === norm_(V_REPROVADO)) {
    if (!g("motivoReprov").trim()) { situacao("Reprovado: escreva o motivo na coluna Motivo da reprovação", V_COR.alerta); return; }
    situacao("Reprovado", V_COR.ruim);
    if (!g("validadoPor")) set("validadoPor", quem || "—");
    if (!g("validadoEm")) set("validadoEm", "'" + agora_());
    return;
  }

  /* decisão apagada ou valor estranho: volta a aguardar */
  situacao(V_AGUARDANDO, V_COR.espera);
  set("validadoPor", ""); set("validadoEm", "");
}

/* Gatilho instalável de edição na planilha de validação (configurarValidacao cria). */
function aoEditarValidacao(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getParent().getId() !== idValidacao_()) return;
  const m = vmapa_(sh);
  const c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
  const mexeu = ["decisao", "motivoReprov"].some(k => m.idx[k] >= 0 && m.idx[k] + 1 >= c1 && m.idx[k] + 1 <= c2);
  if (!mexeu) return;
  let quem = "";
  try { quem = (e.user && e.user.getEmail()) || Session.getActiveUser().getEmail() || ""; } catch (x) {}
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const r1 = Math.max(2, e.range.getRow()), r2 = e.range.getRow() + e.range.getNumRows() - 1;
    for (let r = r1; r <= r2; r++) processarLinha_(sh, m, r, quem);
  } finally { lock.releaseLock(); }
}

/* Rede de segurança: a cada VALIDACAO_VARREDURA_MIN minutos, aplica decisões que o gatilho de edição perdeu
   (colagem em várias linhas, erro temporário ao abrir a planilha da fábrica etc.). */
function varrerValidacao() {
  if (!idValidacao_()) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const sh = abaValidacao_(), m = vmapa_(sh), last = sh.getLastRow();
    if (last < 2) return;
    const vals = sh.getRange(2, 1, last - 1, m.ncols).getDisplayValues();
    vals.forEach((row, i) => {
      const dec = norm_(row[m.idx.decisao]), sit = row[m.idx.situacao];
      const pendente =
        (dec === norm_(V_APROVADO) && !/^(Enviado|Aprovado:)/.test(sit)) ||
        (dec === norm_(V_REPROVADO) && (sit !== "Reprovado" || (m.idx.motivoReprov >= 0 && !row[m.idx.motivoReprov].trim()))) ||
        (/^Enviado/.test(sit) && dec !== norm_(V_APROVADO));
      if (pendente) processarLinha_(sh, m, i + 2, "");
    });
  } finally { lock.releaseLock(); }
}

/* Consulta do site: reenvios dessa order que ainda não foram para a fábrica. */
function buscarNaValidacao_(alvo) {
  if (!idValidacao_()) return [];
  const sh = abaValidacao_(), m = vmapa_(sh), last = sh.getLastRow();
  if (last < 2) return [];
  const g = (row, k) => (m.idx[k] >= 0 ? row[m.idx[k]] : "");
  return sh.getRange(2, 1, last - 1, m.ncols).getDisplayValues()
    .filter(row => orderKey_(g(row, "order")) === alvo && !/^Enviado/.test(g(row, "situacao")))
    .map(row => {
      const sit = g(row, "situacao");
      const status = sit === "Reprovado" ? "Reprovado na validação" + (g(row, "motivoReprov") ? ": " + g(row, "motivoReprov") : "")
        : /^Aprovado:/.test(sit) ? "Aprovado na validação (enviar à fábrica manualmente)"
        : "Aguardando validação";
      return {
        fabrica: g(row, "fabrica"), protocolo: g(row, "protocolo"), ticket: g(row, "ticket"),
        dataEnvio: "", produto: g(row, "produto"), quantidade: g(row, "quantidade"),
        tracking: "", status: status, observacoes: g(row, "obs"), validacao: true,
        recebido: g(row, "recebido")
      };
    });
}

/* Rode UMA vez pelo editor (selecione configurarValidacao e clique em Executar).
   Cria a planilha "Validação de reenvios" (se ainda não existir), protege as colunas,
   e instala o gatilho de edição e a varredura. Pode rodar de novo sem problema. */
function configurarValidacao() {
  const props = PropertiesService.getScriptProperties();
  let id = idValidacao_();
  if (!id) {
    const novo = SpreadsheetApp.create("Validação de reenvios — PagAmerican (interno)");
    id = novo.getId();
    props.setProperty("VALIDACAO_ID", id);
    novo.getSheets()[0].setName(V_ABA);
  }
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(V_ABA) || ss.getSheets()[0];

  /* cabeçalho (só cria se a linha 1 estiver vazia, para não bagunçar uma planilha em uso) */
  if (!sh.getRange(1, 1).getDisplayValue()) {
    sh.getRange(1, 1, 1, VCOLS.length).setValues([VCOLS.map(c => c[1])])
      .setFontWeight("bold").setBackground("#1f3b5c").setFontColor("#ffffff").setWrap(true);
    VCOLS.forEach((c, i) => sh.setColumnWidth(i + 1, c[2]));
    sh.setFrozenRows(1);
  }
  /* linhas de sobra: a lista suspensa e a proteção valem até a última linha que existir agora */
  if (sh.getMaxRows() < 5000) sh.insertRowsAfter(sh.getMaxRows(), 5000 - sh.getMaxRows());
  const m = vmapa_(sh), linhas = sh.getMaxRows() - 1;
  sh.getRange(1, m.idx.decisao + 1).setBackground("#0b6e4f");
  if (m.idx.motivoReprov >= 0) sh.getRange(1, m.idx.motivoReprov + 1).setBackground("#0b6e4f");

  /* lista suspensa na Decisão + cores */
  const colDec = sh.getRange(2, m.idx.decisao + 1, linhas, 1);
  colDec.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([V_APROVADO, V_REPROVADO], true).setAllowInvalid(false)
    .setHelpText("Aprovado = vai para a planilha da fábrica. Reprovado = não vai (escreva o motivo).").build());
  const regras = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(g => g.getColumn() === m.idx.decisao + 1));
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(V_APROVADO).setBackground("#c8e6c9").setBold(true).setRanges([colDec]).build());
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(V_REPROVADO).setBackground("#f8c7c3").setBold(true).setRanges([colDec]).build());
  sh.setConditionalFormatRules(regras);

  /* proteção: só Decisão e Motivo da reprovação ficam editáveis para os editores */
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const prot = sh.protect().setDescription("Só o script altera estas colunas");
  const livres = [sh.getRange(2, m.idx.decisao + 1, linhas, 1)];
  if (m.idx.motivoReprov >= 0) livres.push(sh.getRange(2, m.idx.motivoReprov + 1, linhas, 1));
  prot.setUnprotectedRanges(livres);
  const eu = Session.getEffectiveUser();
  prot.addEditor(eu);
  prot.removeEditors(prot.getEditors().filter(u => u.getEmail() !== eu.getEmail()));
  if (prot.canDomainEdit()) prot.setDomainEdit(false);

  /* gatilhos */
  ScriptApp.getProjectTriggers()
    .filter(t => ["aoEditarValidacao", "varrerValidacao"].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("aoEditarValidacao").forSpreadsheet(id).onEdit().create();
  ScriptApp.newTrigger("varrerValidacao").timeBased().everyMinutes(VALIDACAO_VARREDURA_MIN).create();

  const recebem = destinatarios_(sh);
  console.log("Planilha de validação: " + ss.getUrl());
  console.log("Recebem o aviso: " + (recebem.join(", ") || "ninguém ainda: compartilhe a planilha com a validadora como Editor"));
  console.log("Gatilhos instalados: edição na planilha + varredura a cada " + VALIDACAO_VARREDURA_MIN + " min.");
}

/* Para voltar ao fluxo antigo (direto para a fábrica): rode desligarValidacao. A planilha não é apagada. */
function desligarValidacao() {
  PropertiesService.getScriptProperties().deleteProperty("VALIDACAO_ID");
  ScriptApp.getProjectTriggers()
    .filter(t => ["aoEditarValidacao", "varrerValidacao"].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  console.log("Validação desligada" + (VALIDACAO_ID ? ", mas VALIDACAO_ID está preenchido no código: apague o valor para desligar de vez." : "."));
}
