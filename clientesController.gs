function getClienteLogado() {
  try {
    const user = getUsuarioLogado();

    if (!user || !user.success) {
      return {
        success: false,
        message: "Usuário não autenticado.",
      };
    }

    const role = String(user.role || "")
      .toUpperCase()
      .trim();

    // 🔐 Apenas CLIENTE ou HOST
    if (!["CLIENTE", "HOST"].includes(role)) {
      return {
        success: false,
        message: "Usuário não é cliente.",
      };
    }

    const evento = getEventoAtivoPorEmailCliente(user.email);

    if (!evento) {
      return {
        success: false,
        message: "Nenhum evento ativo encontrado.",
      };
    }

    return {
      success: true,
      tipo: role,
      email: user.email,
      nome: user.name || "",
      idEvento: evento.idEvento,
      nomeEvento: evento.nomeEvento,
    };
  } catch (e) {
    Logger.log("Erro getClienteLogado: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

function garantirUsuarioCliente(email, chefeEmail, senhaEvento) {
  try {
    if (!email) return false;

    const sheet =
      SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName("User");

    if (!sheet) {
      throw new Error("Aba User não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    const emailNormalizado = String(email).toLowerCase().trim();

    // 🔍 Verifica duplicidade
    for (let i = 1; i < data.length; i++) {
      const emailExistente = String(data[i][headerMap.email] || "")
        .toLowerCase()
        .trim();

      if (emailExistente === emailNormalizado) {
        Logger.log("Usuário já existe: " + emailNormalizado);

        return true;
      }
    }

    const now = new Date();

    const username = emailNormalizado.split("@")[0];

    const novaLinha = new Array(Object.keys(headerMap).length).fill("");

    const setCampo = (campo, valor) => {
      if (headerMap[campo] !== undefined) {
        novaLinha[headerMap[campo]] = valor;
      }
    };

    setCampo("id_username", Utilities.getUuid());

    setCampo("username", username);

    setCampo("name", username);

    setCampo("email", emailNormalizado);

    setCampo("phone", "");

    setCampo("role", "CLIENTE");

    setCampo("pin", senhaEvento);

    setCampo("date_creation", now);

    setCampo("chefe_email", chefeEmail || "");

    setCampo("administrative", "NÃO");

    setCampo("managepermissons", "NÃO");

    setCampo("registrationevent", "SIM");

    setCampo("checkin", "NÃO");

    setCampo("rsvp", "SIM");

    setCampo("data_update", now);

    sheet.appendRow(novaLinha);

    return true;
  } catch (e) {
    Logger.log("Erro garantirUsuarioCliente: " + e.message);

    return false;
  }
}

function salvarUsuarioAnfitriao(payload) {
  try {
    const sheetUser =
      SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName("User");

    if (!sheetUser) {
      throw new Error("Aba User não encontrada.");
    }

    const headerMapUser = createHeaderMap(sheetUser);

    const dataUsers = sheetUser.getDataRange().getValues();

    /* =====================================================
       GERA E-MAIL ÚNICO
    ===================================================== */

    const nomeBase = String(payload.hosts || "host")
      .split(" ")[0]
      .toLowerCase()
      .trim()
      .replace(/\s+/g, "");

    let contador = 0;

    let novoEmail = "";

    while (true) {
      novoEmail =
        contador === 0
          ? `${nomeBase}@anfitriao.com`
          : `${nomeBase}${contador}@anfitriao.com`;

      const existe = dataUsers.some((row) => {
        const email = String(row[headerMapUser.email] || "")
          .toLowerCase()
          .trim();

        return email === novoEmail;
      });

      if (!existe) break;

      contador++;
    }

    /* =====================================================
       MONTA LINHA
    ===================================================== */

    const novaLinha = new Array(Object.keys(headerMapUser).length).fill("");

    const setCampo = (campo, valor) => {
      if (headerMapUser[campo] !== undefined) {
        novaLinha[headerMapUser[campo]] = valor;
      }
    };

    setCampo("id_username", Utilities.getUuid());

    setCampo("username", nomeBase);

    setCampo("name", payload.hosts);

    setCampo("email", novoEmail);

    setCampo("role", "CLIENTE");

    setCampo("pin", payload.hostpassword);

    setCampo("date_creation", new Date());

    setCampo("chefe_email", payload.cadastrado_por || "");

    setCampo("registration", "ATIVO");

    sheetUser.appendRow(novaLinha);

    /* =====================================================
       ATUALIZA EVENTS
    ===================================================== */

    atualizarEmailAnfitriaoEvento_(payload.idevento, novoEmail);

    return {
      success: true,
      email: novoEmail,
    };
  } catch (e) {
    Logger.log("Erro salvarUsuarioAnfitriao: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

function atualizarEmailAnfitriaoEvento_(idevento, email) {
  try {
    if (!idevento) return;

    const sheet =
      SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName("Events");

    if (!sheet) return;

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const evt = String(data[i][headerMap.idevento]);

      if (evt === String(idevento)) {
        if (headerMap.emailanfitriao !== undefined) {
          sheet.getRange(i + 1, headerMap.emailanfitriao + 1).setValue(email);

          Logger.log("E-mail anfitrião atualizado.");
        }

        return;
      }
    }
  } catch (e) {
    Logger.log("Erro atualizarEmailAnfitriaoEvento_: " + e.message);
  }
}

function obterDadosEListaCompleta(emailUsuario) {
  try {
    const evento = obterDadosPorEmail(emailUsuario);

    if (!evento || !evento.idEvento) {
      return {
        success: false,
        mensagem: "Evento não localizado.",
      };
    }

    const lista = listarConvidadosAnfitriao(evento.idEvento);

    return {
      success: true,

      idEvento: evento.idEvento,

      nomeEvento: evento.nome,

      convidados: lista || [],
    };
  } catch (e) {
    Logger.log("Erro obterDadosEListaCompleta: " + e.message);

    return {
      success: false,
      mensagem: e.message,
    };
  }
}
