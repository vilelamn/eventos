/* =========================================================
   RSVP.GS
========================================================= */

/* =========================================================
   SALVAR CONVIDADO
========================================================= */
function salvarConvidadoStep1(dados, idEvento) {
  try {
    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const idGuest = "GUEST-" + new Date().getTime();

    const qrCodeUrl =
      "https://quickchart.io/qr?text=" +
      encodeURIComponent(idGuest) +
      "&size=200";

    const novaLinha = Array(sheet.getLastColumn()).fill("");

    const campos = {
      idevento: idEvento,
      idguest: idGuest,
      guestname: String(dados.nome || "").trim(),
      typeguest: "INDIVIDUAL",
      guestemail: normalizarEmail(dados.email),
      phone: "'" + String(dados.telefone || "").trim(),
      presence: "SIM",
      adult_companions: 0,
      qrcode: qrCodeUrl,
      status: "CONFIRMADO",
      checkin: "NAO",
      checkindate: "",
      photourl: "PENDENTE",
    };

    Object.keys(campos).forEach((campo) => {
      if (headerMap[campo] !== undefined) {
        novaLinha[headerMap[campo]] = campos[campo];
      }
    });

    sheet.appendRow(novaLinha);

    try {
      if (typeof enviarEmailAgradecimento === "function") {
        enviarEmailAgradecimento(dados.nome, dados.email, qrCodeUrl);
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
  }
}

/* =========================================================
   REGISTRAR CHECK-IN
========================================================= */
function registrarCheckin(idConvidado, idEvento) {
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

        if (headerMap.status !== undefined) {
          sheet.getRange(linha, headerMap.status + 1).setValue("PRESENTE");
        }

        if (headerMap.checkin !== undefined) {
          sheet.getRange(linha, headerMap.checkin + 1).setValue("SIM");
        }

        if (headerMap.checkindate !== undefined) {
          sheet.getRange(linha, headerMap.checkindate + 1).setValue(new Date());
        }

        return {
          success: true,
          mensagem: "Check-in realizado com sucesso.",
          guestname: dados[i][headerMap.guestname],
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
  }
}

/* =========================================================
   VERIFICAR CHECK-IN
========================================================= */
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

/* =========================================================
   ATUALIZAR STATUS
========================================================= */
function atualizarStatusConvidado(idConvidado, novoStatus) {
  try {
    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const headerMap = createHeaderMap(sheet);

    const dados = sheet.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const guest = String(dados[i][headerMap.idguest] || "").trim();

      if (guest === String(idConvidado).trim()) {
        sheet
          .getRange(i + 1, headerMap.status + 1)
          .setValue(String(novoStatus).toUpperCase().trim());

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

/* =========================================================
   PROCESSAR CHECK-IN COM FOTO
========================================================= */
function processarCheckinComFoto(idConvidado, idEvento, fotoUrl = "") {
  try {
    const resultado = registrarCheckin(idConvidado, idEvento);

    if (!resultado.success) {
      return resultado;
    }

    if (fotoUrl) {
      const sheet = getSheetRSVPs();

      const headerMap = createHeaderMap(sheet);

      const dados = sheet.getDataRange().getValues();

      for (let i = 1; i < dados.length; i++) {
        const evt = String(dados[i][headerMap.idevento] || "").trim();

        const guest = String(dados[i][headerMap.idguest] || "").trim();

        if (
          evt === String(idEvento).trim() &&
          guest === String(idConvidado).trim()
        ) {
          if (headerMap.photourl !== undefined) {
            sheet.getRange(i + 1, headerMap.photourl + 1).setValue(fotoUrl);
          }

          break;
        }
      }
    }

    return {
      success: true,
      mensagem: "Check-in realizado.",
    };
  } catch (e) {
    Logger.log("Erro processarCheckinComFoto: " + e.message);

    return {
      success: false,
      mensagem: e.message,
    };
  }
}
