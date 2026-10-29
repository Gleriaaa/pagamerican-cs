/* ============================================================
   Reenvio · planilhas das fábricas (EagleLabs, ShipOffers, RedRock)
   ------------------------------------------------------------
   Web App separado do script da aba "Reenvios" (reenvio.gs).
   - GET  ?chave=...&order=183156  → tracking, status e data do reenvio
                                     (nunca devolve e-mail nem endereço)
   - POST {chave, acao:"adicionar", ...} → acrescenta o reenvio na
                                     planilha da fábrica certa
   Como publicar: ver apps-script/README-reenvio-fabricas.md
   ============================================================ */

const CHAVE = "pa-trk-sbkaevI2NaUoIYQ1nsHi";

/* Uma entrada por fábrica: o ID é o trecho do link da planilha entre /d/ e /edit.
   "aba" vazio = primeira aba da planilha. */
const FABRICAS = {
  eaglelabs:  { nome: "EagleLabs",  id: "COLE_O_ID_DA_PLANILHA_EAGLELABS",  aba: "" },
  shipoffers: { nome: "ShipOffers", id: "COLE_O_ID_DA_PLANILHA_SHIPOFFERS", aba: "" },
  redrock:    { nome: "RedRock",    id: "COLE_O_ID_DA_PLANILHA_REDROCK",    aba: "" }
};

const STATUS_INICIAL = "Enviado à fábrica";

/* Cabeçalhos procurados (sem acento, minúsculo, só letras e números). */
const COLS = {
  order:      ["ordernumber", "order", "pedido"],
  produto:    ["produto", "product"],
  quantidade: ["quantidade", "qtd", "quantity"],
  endereco:   ["enderecodeenvio", "shippingaddress", "endereco", "address"],
  email:      ["customeremail", "email"],
  protocolo:  ["protocolo"],
  ticket:     ["ticket"],
  dataEnvio:  ["datadoenvioafabrica", "dataenvio", "data"],
  tracking:   ["trackingdoreenvio", "tracking", "rastreio"],
  status:     ["status"],
  obs:        ["observacoes", "obs"]
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
    return json_({ ok: true, fabrica: f.nome });
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
