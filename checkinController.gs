function salvarConvidadoStep1(dados, idEvento) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!idEvento) {
      throw new Error("Evento não informado.");
    }

    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    const email = normalizarEmail(dados.email);
    const telefone = String(dados.telefone || "").trim();

    // ==========================================
    // EVITA DUPLICIDADE
    // ==========================================

    for (let i = 1; i < data.length; i++) {
      const evt = String(data[i][headerMap.idevento] || "").trim();

      const emailExistente = normalizarEmail(data[i][headerMap.guestemail]);

      if (
        evt === String(idEvento).trim() &&
        emailExistente &&
        emailExistente === email
      ) {
        return {
          success: false,
          mensagem: "Este convidado já foi cadastrado.",
        };
      }
    }

    const idGuest = gerarUUID();

    const qrCodeUrl = `https://quickchart.io/qr?text=${encodeURIComponent(idGuest)}&size=200`;

    const novaLinha = Array(sheet.getLastColumn()).fill("");

    const campos = {
      idevento: idEvento,

      idguest: idGuest,

      guestname: String(dados.nome || "").trim(),

      typeguest: "INDIVIDUAL",

      guestemail: email,

      phone: telefone,

      presence: "SIM",

      adult_companions: 0,

      qrcode: qrCodeUrl,

      status: "CONFIRMADO",

      checkin: "NAO",

      checkindate: "",

      photourl: "PENDENTE",

      categoria: String(dados.categoria || "CONVIDADO").trim(),

      date_creation: new Date(),
    };

    Object.keys(campos).forEach((campo) => {
      if (headerMap[campo] !== undefined) {
        novaLinha[headerMap[campo]] = campos[campo];
      }
    });

    sheet.appendRow(novaLinha);

    // ==========================================
    // ENVIO DE EMAIL
    // ==========================================

    try {
      if (typeof enviarEmailAgradecimento === "function" && email) {
        enviarEmailAgradecimento(dados.nome, email, qrCodeUrl);
      }
    } catch (e) {
      Logger.log("Erro enviarEmailAgradecimento: " + e.message);
    }

    return {
      success: true,
      idguest: idGuest,
      qrcode: qrCodeUrl,
    };
  } catch (e) {
    Logger.log("Erro salvarConvidadoStep1: " + e.message);

    return {
      success: false,
      mensagem: e.message,
    };
  } finally {
    lock.releaseLock();
  }
}

function registrarCheckin(idConvidado, idEvento) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const dados = sheet.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const evt = String(dados[i][headerMap.idevento] || "").trim();

      const guest = String(dados[i][headerMap.idguest] || "").trim();

      if (
        evt === String(idEvento).trim() &&
        guest === String(idConvidado).trim()
      ) {
        const jaFezCheckin =
          String(dados[i][headerMap.checkin] || "")
            .toUpperCase()
            .trim() === "SIM";

        if (jaFezCheckin) {
          return {
            success: false,
            mensagem: "Check-in já realizado.",
          };
        }

        const linha = i + 1;

        const updates = [];

        if (headerMap.status !== undefined) {
          updates.push({
            col: headerMap.status + 1,
            valor: "PRESENTE",
          });
        }

        if (headerMap.checkin !== undefined) {
          updates.push({
            col: headerMap.checkin + 1,
            valor: "SIM",
          });
        }

        if (headerMap.checkindate !== undefined) {
          updates.push({
            col: headerMap.checkindate + 1,
            valor: new Date(),
          });
        }

        updates.forEach((u) => {
          sheet.getRange(linha, u.col).setValue(u.valor);
        });

        return {
          success: true,
          mensagem: "Check-in realizado com sucesso.",
          guestname: dados[i][headerMap.guestname] || "",
        };
      }
    }

    throw new Error("Convidado não encontrado.");
  } catch (e) {
    Logger.log("Erro registrarCheckin: " + e.message);

    return {
      success: false,
      mensagem: e.message,
    };
  } finally {
    lock.releaseLock();
  }
}

function atualizarStatusConvidado(idConvidado, novoStatus) {
  try {
    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const dados = sheet.getDataRange().getValues();

    const statusNormalizado = String(novoStatus || "")
      .toUpperCase()
      .trim();

    for (let i = 1; i < dados.length; i++) {
      const guest = String(dados[i][headerMap.idguest] || "").trim();

      if (guest === String(idConvidado).trim()) {
        sheet.getRange(i + 1, headerMap.status + 1).setValue(statusNormalizado);

        return {
          success: true,
          message: "Status atualizado.",
        };
      }
    }

    throw new Error("Convidado não encontrado.");
  } catch (e) {
    Logger.log("Erro atualizarStatusConvidado: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

function verificarCheckin(idConvidado, idEvento) {
  try {
    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const dados = sheet.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const evt = String(dados[i][headerMap.idevento] || "").trim();

      const guest = String(dados[i][headerMap.idguest] || "").trim();

      if (
        evt === String(idEvento).trim() &&
        guest === String(idConvidado).trim()
      ) {
        return {
          success: true,

          checkin:
            String(dados[i][headerMap.checkin] || "")
              .toUpperCase()
              .trim() === "SIM",

          status: dados[i][headerMap.status] || "",

          guestname: dados[i][headerMap.guestname] || "",

          checkindate:
            dados[i][headerMap.checkindate] instanceof Date
              ? formatarData(
                  dados[i][headerMap.checkindate],
                  "dd/MM/yyyy HH:mm",
                )
              : "",
        };
      }
    }

    return {
      success: false,
      mensagem: "Convidado não encontrado.",
    };
  } catch (e) {
    Logger.log("Erro verificarCheckin: " + e.message);

    return {
      success: false,
      mensagem: e.message,
    };
  }
}
