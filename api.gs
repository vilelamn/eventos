/* =========================================================
   API.GS
   CENTRAL DE ROTAS DA API + HTML
========================================================= */

/* =========================================================
   WEBAPP
========================================================= */

function doGet(e) {
  try {
    // =====================================================
    // API
    // =====================================================
    if (e && e.parameter && e.parameter.acao) {
      return processarRequisicao(e);
    }

    // =====================================================
    // HTML
    // =====================================================
    const pagina = e && e.parameter && e.parameter.p ? e.parameter.p : "login";

    return carregarTemplateHtml(pagina);
  } catch (erro) {
    Logger.log("ERRO doGet: " + (erro.stack || erro.message || erro));

    return HtmlService.createHtmlOutput("<h1>Erro interno</h1>");
  }
}

function doPost(e) {
  return processarRequisicao(e);
}

/* =========================================================
   PROCESSADOR CENTRAL
========================================================= */

function processarRequisicao(e) {
  try {
    if (!e) {
      return respostaJSON(false, "Requisição inválida.");
    }

    let acao = "";
    let dados = {};

    /* =====================================================
       POST JSON
    ===================================================== */

    if (e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);

        acao = String(body.acao || "").trim();

        dados = body.dados || {};
      } catch (jsonError) {
        return respostaJSON(false, "JSON inválido.");
      }
    } else if (e.parameter) {

    /* =====================================================
       GET PARAMS
    ===================================================== */
      acao = String(e.parameter.acao || "").trim();

      dados = e.parameter;
    }

    /* =====================================================
       VALIDAÇÃO
    ===================================================== */

    if (!acao) {
      return respostaJSON(false, "Ação não informada.");
    }

    Logger.log("AÇÃO API: " + acao);

    /* =====================================================
       ROTAS PÚBLICAS
    ===================================================== */

    const rotasPublicas = ["login", "salvarConvidadoStep1", "recuperar"];

    /* =====================================================
       AUTENTICAÇÃO
    ===================================================== */

    if (!rotasPublicas.includes(acao)) {
      const user = getUsuarioLogado();

      if (!user || !user.success) {
        return respostaJSON(false, "Usuário não autenticado.");
      }
    }

    /* =====================================================
       SWITCH ROTAS
    ===================================================== */

    switch (acao) {
      /* =================================================
         LOGIN
      ================================================= */

      case "login":
        if (!dados.email || !dados.pin) {
          return respostaJSON(false, "E-mail e senha obrigatórios.");
        }

        return respostaObjeto(realizarLogin(dados.email, dados.pin));

      /* =================================================
         DASHBOARD
      ================================================= */

      case "getDashboardStats":
        return respostaObjeto(getDashboardStats());

      /* =================================================
         EVENTOS
      ================================================= */

      case "listarEventos":

      case "listarEventosTabela":
        return respostaObjeto(listarEventosTabela(dados.email, dados.role));

      case "getEventoById":
        if (!dados.id) {
          return respostaJSON(false, "ID do evento obrigatório.");
        }

        return respostaObjeto(getEventoById(dados.id, dados.email, dados.role));

      case "processarSalvarEvento":
        return respostaObjeto(processarSalvarEvento(dados));

      case "excluirEventoNoServidor":
        if (!dados.id) {
          return respostaJSON(false, "ID do evento obrigatório.");
        }

        return respostaObjeto(excluirEventoNoServidor(dados.id));

      /* =================================================
         CONVIDADOS
      ================================================= */

      case "salvarConvidadoStep1":
        if (!dados.idEvento) {
          return respostaJSON(false, "Evento obrigatório.");
        }

        return respostaObjeto(salvarConvidadoStep1(dados, dados.idEvento));

      case "listarConvidadosRSVP":
        if (!dados.idEvento) {
          return respostaJSON(false, "Evento obrigatório.");
        }

        return respostaObjeto(listarConvidadosRSVP(dados.idEvento));

      case "buscarDadosConvidado":
        if (!dados.idConvidado || !dados.idEvento) {
          return respostaJSON(false, "Parâmetros inválidos.");
        }

        return respostaObjeto(
          buscarDadosConvidado(dados.idConvidado, dados.idEvento),
        );

      /* =================================================
         CHECKIN
      ================================================= */

      case "verificarCheckin":
        if (!dados.idConvidado || !dados.idEvento) {
          return respostaJSON(false, "Parâmetros inválidos.");
        }

        return respostaObjeto(
          verificarCheckin(dados.idConvidado, dados.idEvento),
        );

      case "processarCheckinComFoto":
        if (!dados.idConvidado || !dados.idEvento) {
          return respostaJSON(false, "Parâmetros inválidos.");
        }

        return respostaObjeto(
          processarCheckinComFoto(dados.idConvidado, dados.idEvento),
        );

      /* =================================================
         CERIMONIALISTAS
      ================================================= */

      case "getCerimonialistas":
        return respostaObjeto({
          success: true,

          data: getCerimonialistas(),
        });

      /* =================================================
         CLIENTE
      ================================================= */

      case "getClienteLogado":
        return respostaObjeto(getClienteLogado());

      case "obterDadosEListaCompleta":
        if (!dados.email) {
          return respostaJSON(false, "E-mail obrigatório.");
        }

        return respostaObjeto(obterDadosEListaCompleta(dados.email));

      /* =================================================
         RECUPERAÇÃO
      ================================================= */

      case "recuperar":
        if (!dados.email) {
          return respostaJSON(false, "E-mail obrigatório.");
        }

        return respostaObjeto({
          success: true,

          message: enviarRecuperacaoSenha(dados.email),
        });

      /* =================================================
         CONTRATOS
      ================================================= */

      case "salvarContrato":
        return respostaObjeto(salvarContratoSaaS(dados));

      /* =================================================
         DEFAULT
      ================================================= */

      default:
        return respostaJSON(false, "Ação inválida: " + acao);
    }
  } catch (erro) {
    Logger.log("ERRO API: " + (erro.stack || erro.message || erro));

    return respostaJSON(
      false,

      erro.message || "Erro interno.",
    );
  }
}

/* =========================================================
   RESPOSTAS JSON
========================================================= */

function respostaJSON(success, message, extra = {}) {
  return ContentService.createTextOutput(
    JSON.stringify({
      success,
      message,

      ...extra,
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}

function respostaObjeto(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj || {})).setMimeType(
    ContentService.MimeType.JSON,
  );
}
