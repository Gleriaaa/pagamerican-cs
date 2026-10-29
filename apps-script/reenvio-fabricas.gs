/* ============================================================
   Reenvio · planilhas das fábricas (EagleLabs, ShipOffers, RedRock)
   ------------------------------------------------------------
   Web App separado do script da aba "Reenvios" (reenvio.gs).
   - GET  ?chave=...&order=183156  → tracking, status e data do reenvio
                                     (nunca devolve e-mail nem endereço)
   - POST {chave, acao:"adicionar", ...} → acrescenta o reenvio na
                                     planilha da fábrica certa e avisa a
                                     fábrica por e-mail
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

/* Cabeçalhos procurados, em português ou inglês (sem acento, minúsculo, só letras e números). */
const COLS = {
  order:      ["ordernumber", "order", "pedido"],
  produto:    ["produto", "product"],
  quantidade: ["quantidade", "qtd", "quantity"],
  endereco:   ["enderecodeenvio", "shippingaddress", "endereco", "address"],
  email:      ["customeremail", "email"],
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
    const f = fabrica_(d.fabrica);
    if (!f || !f.id || f.id.indexOf("COLE_") === 0)
      return json_({ ok: false, semPlanilha: true, erro: "A fábrica " + (d.fabrica || "?") + " ainda não tem planilha de acompanhamento." });

    lock.waitLock(20000);
    const sh = aba_(f), m = mapa_(sh), last = sh.getLastRow();

    /* não duplica o mesmo protocolo */
    if (d.protocolo && m.idx.protocolo >= 0 && last > m.linha) {
      const protos = sh.getRange(m.linha + 1, m.idx.protocolo + 1, last - m.linha, 1).getDisplayValues().map(r => r[0]);
      if (protos.indexOf(d.protocolo) >= 0) return json_({ ok: true, fabrica: f.nome, jaExistia: true });
    }

    const row = new Array(m.ncols).fill("");
    const put = (c, v) => { if (m.idx[c] >= 0) row[m.idx[c]] = v == null ? "" : String(v); };
    put("order", d.orderNumber);
    put("produto", d.produto);
    put("quantidade", d.quantidade);
    put("endereco", d.endereco);
    put("email", d.email);
    put("protocolo", d.protocolo);
    put("ticket", d.ticket);
    put("dataEnvio", Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy"));
    put("status", STATUS_INICIAL);
    sh.appendRow(row);
    let avisados = 0;
    try { avisados = avisarNovo_(f, sh, d); } catch (err) { console.error("Aviso por e-mail falhou: " + err); }
    return json_({ ok: true, fabrica: f.nome, avisados: avisados });
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
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

function enviar_(para, assunto, texto, html) {
  if (!para.length) return 0;
  const opt = { htmlBody: html, name: ASSINATURA };
  if (AVISO_COPIA.length) opt.cc = AVISO_COPIA.join(",");
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
