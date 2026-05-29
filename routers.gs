function getPaginasPermitidas() {
  return [
    "index",
    "login",
    "dashboard",
    "eventos",
    "checkin",
    "confirmacao",
    "clientes",
    "usuarios",
    "relatorios",
  ];
}

function getPaginasPublicas() {
  return ["index", "login", "confirmacao"];
}

function paginaPermitida(page) {
  return getPaginasPermitidas().includes(page);
}

function carregarTemplateHtml(pagina) {
  try {
    const page = String(pagina || "index")
      .trim()
      .toLowerCase();

    // ==========================================
    // VALIDA PÁGINA
    // ==========================================

    if (!paginaPermitida(page)) {
      return HtmlService.createHtmlOutput("<h2>Página não permitida.</h2>");
    }

    // ==========================================
    // CONTROLE LOGIN
    // ==========================================

    const paginasPublicas = getPaginasPublicas();

    const user = getUsuarioLogado();

    const precisaLogin = !paginasPublicas.includes(page);

    if (precisaLogin && (!user || !user.success)) {
      return HtmlService.createHtmlOutput("<h2>Acesso não autorizado.</h2>");
    }

    // ==========================================
    // TEMPLATE
    // ==========================================

    const template = HtmlService.createTemplateFromFile(page);

    template.user = user || {};

    template.appName = "MeuEvento";

    template.currentPage = page;

    // ==========================================
    // RETORNO
    // ==========================================

    return template
      .evaluate()
      .setTitle("MeuEvento")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } catch (e) {
    Logger.log("Erro carregarTemplateHtml: " + e.message);

    return HtmlService.createHtmlOutput(`
      <div style="padding:20px;font-family:Arial">
        <h2>Erro interno</h2>
        <p>${e.message}</p>
      </div>
    `);
  }
}

function getPage(pageName) {
  try {
    const page = String(pageName || "index")
      .trim()
      .toLowerCase();

    // ==========================================
    // VALIDA PÁGINA
    // ==========================================

    if (!paginaPermitida(page)) {
      return `
        <div style="padding:20px">
          Página não permitida.
        </div>
      `;
    }

    const user = getUsuarioLogado() || {};

    const paginasPublicas = getPaginasPublicas();

    const precisaLogin = !paginasPublicas.includes(page);

    if (precisaLogin && !user.success) {
      return `
        <div style="padding:20px">
          Acesso não autorizado.
        </div>
      `;
    }

    // ==========================================
    // TEMPLATE
    // ==========================================

    const template = HtmlService.createTemplateFromFile(page);

    template.user = user;

    template.appName = "MeuEvento";

    template.currentPage = page;

    return template.evaluate().getContent();
  } catch (e) {
    Logger.log("Erro getPage: " + e.message);

    return `
      <div style="padding:20px;font-family:Arial">
        <h2>Erro ao carregar página</h2>
        <p>${e.message}</p>
      </div>
    `;
  }
}
