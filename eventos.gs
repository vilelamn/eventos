function listarEventosTabela(usuarioLogadoEmail, roleLogado) {
  try {
    const roleBusca = normalizarRole(roleLogado);
    const emailBusca = normalizarEmail(usuarioLogadoEmail);

    const sheet = getSheetEventos();

    if (!sheet) return [];

    const headerMap = createHeaderMap(sheet);

    if (
      headerMap.idevento === undefined ||
      headerMap.cadastrado_por === undefined ||
      headerMap.eventdate === undefined ||
      headerMap.eventhour === undefined
    ) {
      throw new Error("Colunas obrigatórias ausentes.");
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return [];

    const isAdmin = ["ADMIN", "ADMINISTRADOR"].includes(roleBusca);

    const emailsPermitidos = isAdmin
      ? []
      : getEmailsDoGrupoCerimonial_(emailBusca, roleBusca);

    return data
      .slice(1)
      .filter((row) => {
        if (isAdmin) return true;

        const dono = normalizarEmail(row[headerMap.cadastrado_por]);

        return emailsPermitidos.includes(dono);
      })
      .map((row) => {
        const dataEvento = row[headerMap.eventdate];
        const horaEvento = row[headerMap.eventhour];

        return {
          idevento: String(row[headerMap.idevento] || ""),

          eventyname: String(row[headerMap.eventyname] || ""),

          eventdate:
            dataEvento instanceof Date ? formatarData(dataEvento) : dataEvento,

          eventhour:
            horaEvento instanceof Date
              ? formatarData(horaEvento, "HH:mm")
              : horaEvento,

          eventlocation: String(row[headerMap.eventlocation] || ""),

          esteemed_guests: Number(row[headerMap.esteemed_guests] || 0),

          status: String(row[headerMap.status] || "ATIVO")
            .toUpperCase()
            .trim(),

          url: String(row[headerMap.url] || ""),
        };
      });
  } catch (e) {
    Logger.log("Erro listarEventosTabela: " + e.message);

    return [];
  }
}

function salvarEvento(payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    validarPermissaoEvento(payload.idevento);

    const sheet = getSheetEventos();

    if (!sheet) {
      throw new Error("Aba Events não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    payload.eventyname = String(payload.eventyname || "").trim();

    payload.eventlocation = String(payload.eventlocation || "").trim();

    payload.status = String(payload.status || "ATIVO")
      .toUpperCase()
      .trim();

    payload.hosts = String(payload.hosts || "").trim();

    payload.emailanfitriao = normalizarEmail(payload.emailanfitriao);

    if (!payload.eventyname) {
      throw new Error("Nome do evento obrigatório.");
    }

    let rowIndex = -1;
    let existingRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][headerMap.idevento]) === String(payload.idevento)) {
        rowIndex = i + 1;
        existingRow = data[i];
        break;
      }
    }

    const isNovo = rowIndex === -1;

    if (isNovo) {
      payload.idevento = gerarUUID();
    }

    const baseUrl = getWebAppUrl();

    const urlEvento = `${baseUrl}?p=confirmacao&id=${payload.idevento}`;

    const hostNome = String(payload.hosts || "EV");

    const prefixo = hostNome.substring(0, 2).toUpperCase();

    const dataLimpa = String(payload.eventdate || "").replace(/-/g, "");

    const senhaGerada = payload.hostpassword || `${prefixo}${dataLimpa}`;

    const totalColunas = sheet.getLastColumn();

    const novaLinha = existingRow
      ? [...existingRow]
      : Array(totalColunas).fill("");

    const campos = {
      idevento: payload.idevento,

      eventyname: payload.eventyname,

      eventlocation: payload.eventlocation,

      eventdate: payload.eventdate,

      eventhour: payload.eventhour,

      esteemed_guests: Number(payload.esteemed_guests || 0),

      cadastrado_por: normalizarEmail(payload.cadastrado_por),

      status: payload.status,

      hosts: payload.hosts,

      url: urlEvento,

      hostpassword: senhaGerada,

      emailanfitriao: payload.emailanfitriao,
    };

    Object.keys(campos).forEach((key) => {
      if (headerMap[key] !== undefined) {
        novaLinha[headerMap[key]] = campos[key];
      }
    });

    if (isNovo) {
      sheet.appendRow(novaLinha);
    } else {
      sheet.getRange(rowIndex, 1, 1, novaLinha.length).setValues([novaLinha]);
    }

    if (typeof garantirUsuarioCliente === "function") {
      garantirUsuarioCliente(
        payload.emailanfitriao,
        payload.cadastrado_por,
        senhaGerada,
      );
    }

    return {
      success: true,

      idevento: payload.idevento,

      url: urlEvento,
    };
  } catch (e) {
    Logger.log("Erro salvarEvento: " + e.message);

    return {
      success: false,

      message: e.message,
    };
  } finally {
    lock.releaseLock();
  }
}

function processarSalvarEvento(payload) {
  const user = getUsuarioLogado();

  if (!user || !user.success) {
    throw new Error("Usuário não autenticado.");
  }

  payload.cadastrado_por = user.email;

  payload.ceremonialist = user.email;

  payload.status = payload.status || "ATIVO";

  return salvarEvento(payload);
}

function excluirEventoNoServidor(idevento) {
  validarPermissaoEvento(idevento);

  const sheet = getSheetEventos();

  if (!sheet) {
    throw new Error("Aba Events não encontrada.");
  }

  const headerMap = createHeaderMap(sheet);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headerMap.idevento]) === String(idevento)) {
      sheet.deleteRow(i + 1);

      return {
        success: true,

        message: "Evento removido com sucesso.",
      };
    }
  }

  return {
    success: false,

    message: "Evento não encontrado.",
  };
}

function listarEventosParaCheckin() {
  try {
    const user = getUsuarioLogado();

    if (!user || !user.success) {
      return [];
    }

    const role = normalizarRole(user.role);

    const email = normalizarEmail(user.email);

    const isAdmin = ["ADMIN", "ADMINISTRADOR"].includes(role);

    const emailsPermitidos = isAdmin
      ? []
      : getEmailsDoGrupoCerimonial_(email, role);

    const sheet = getSheetEventos();

    if (!sheet) return [];

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    return data
      .slice(1)
      .filter((row) => {
        const status = String(row[headerMap.status] || "")
          .toUpperCase()
          .trim();

        if (status === "CANCELADO") {
          return false;
        }

        if (isAdmin) return true;

        const dono = normalizarEmail(row[headerMap.cadastrado_por]);

        return emailsPermitidos.includes(dono);
      })
      .map((row) => ({
        idevento: String(row[headerMap.idevento] || ""),

        eventyname: String(row[headerMap.eventyname] || "Sem Nome"),

        eventdate:
          row[headerMap.eventdate] instanceof Date
            ? formatarData(row[headerMap.eventdate])
            : row[headerMap.eventdate],

        eventhour:
          row[headerMap.eventhour] instanceof Date
            ? formatarData(row[headerMap.eventhour], "HH:mm")
            : row[headerMap.eventhour],

        eventlocation: String(row[headerMap.eventlocation] || ""),
      }));
  } catch (e) {
    Logger.log("Erro listarEventosParaCheckin: " + e.message);

    return [];
  }
}

function validarPermissaoEvento(idevento) {
  const user = getUsuarioLogado();

  if (!user || !user.email || !user.role) {
    throw new Error("Usuário não autenticado.");
  }

  const role = normalizarRole(user.role);

  if (["ADMIN", "ADMINISTRADOR"].includes(role)) {
    return;
  }

  if (!["CERIMONIALISTA", "CLIENTE"].includes(role)) {
    throw new Error("Você não tem permissão.");
  }

  if (!idevento) {
    return;
  }

  const evento = getEventoById(idevento, user.email, role);

  if (!evento) {
    throw new Error("Evento não encontrado.");
  }
}

function getEventoById(id, userEmail, userRole) {
  try {
    const role = String(userRole || "")
      .toUpperCase()
      .trim();

    const email = String(userEmail || "")
      .toLowerCase()
      .trim();

    const sheet = getSheetEventos();

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    const colId = headerMap.idevento;

    const isAdmin = role === "ADMIN" || role === "ADMINISTRADOR";

    const emailsPermitidos = isAdmin
      ? []
      : getEmailsDoGrupoCerimonial_(email, role);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colId]) === String(id)) {
        const cadastradoPor = String(data[i][headerMap.cadastrado_por] || "")
          .toLowerCase()
          .trim();

        if (!isAdmin && !emailsPermitidos.includes(cadastradoPor)) {
          throw new Error("Acesso negado ao evento.");
        }

        const obj = {};

        for (const key in headerMap) {
          let val = data[i][headerMap[key]];

          if (key === "eventdate" && val instanceof Date) {
            val = Utilities.formatDate(
              val,
              Session.getScriptTimeZone(),
              "yyyy-MM-dd",
            );
          }

          if (key === "eventhour" && val instanceof Date) {
            val = Utilities.formatDate(
              val,
              Session.getScriptTimeZone(),
              "HH:mm",
            );
          }

          obj[key] = val === undefined || val === null ? "" : val;
        }

        return obj;
      }
    }

    return null;
  } catch (e) {
    Logger.log("Erro em getEventoById: " + e.message);

    throw new Error(e.message);
  }
}

function obterDadosPorEmail(emailUsuario) {
  if (!emailUsuario) return null;

  const ss = SpreadsheetApp.openById(ID_DA_PLANILHA);

  const sheet = ss.getSheetByName("Events");

  if (!sheet) return null;

  const headerMap = createHeaderMap(sheet);

  const dados = sheet.getDataRange().getValues();

  const emailBusca = String(emailUsuario).toLowerCase().trim();

  const colEmail = headerMap.emailanfitriao;

  for (let i = 1; i < dados.length; i++) {
    const emailPlanilha = String(dados[i][colEmail] || "")
      .toLowerCase()
      .trim();

    if (emailPlanilha === emailBusca) {
      return {
        idEvento: dados[i][headerMap.idevento],

        nome: dados[i][headerMap.eventyname],
      };
    }
  }

  return null;
}

function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(ID_DA_PLANILHA);

    const sheetEvents = ss.getSheetByName(ABA_EVENTOS);

    const sheetUser = ss.getSheetByName(ABA_USUARIOS);

    const sheetRSVP = ss.getSheetByName(ABA_RSVPS);

    let countEventos = 0;
    let countCerimonialistas = 0;

    let totalConvidados = 0;

    if (sheetEvents) {
      const dataEvents = sheetEvents.getDataRange().getValues();

      const headerMapEvents = createHeaderMap(sheetEvents);

      const colDono = headerMapEvents.cadastrado_por;

      for (let i = 1; i < dataEvents.length; i++) {
        const emailCriador = String(dataEvents[i][colDono] || "").trim();

        if (emailCriador !== "") {
          countEventos++;
        }
      }
    }

    if (sheetUser) {
      const dataUsers = sheetUser.getDataRange().getValues();

      const headerMapUsers = createHeaderMap(sheetUser);

      const colRole = headerMapUsers.role;

      for (let i = 1; i < dataUsers.length; i++) {
        const role = String(dataUsers[i][colRole] || "")
          .toUpperCase()
          .trim();

        if (role === "CERIMONIALISTA") {
          countCerimonialistas++;
        }
      }
    }

    if (sheetRSVP) {
      totalConvidados = Math.max(sheetRSVP.getLastRow() - 1, 0);
    }

    return {
      success: true,

      eventosAtivos: countEventos,

      cerimonialistas: countCerimonialistas,

      totalConvidados,
    };
  } catch (e) {
    Logger.log("Erro getDashboardStats: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}
