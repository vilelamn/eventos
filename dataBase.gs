/* =========================================================
   SPREADSHEET PRINCIPAL
========================================================= */
function getSpreadsheet() {
  try {
    if (!ID_DA_PLANILHA) {
      throw new Error("ID_DA_PLANILHA não definido.");
    }

    return SpreadsheetApp.openById(ID_DA_PLANILHA);
  } catch (e) {
    Logger.log("Erro getSpreadsheet: " + e.message);

    throw e;
  }
}

/* =========================================================
   FUNÇÃO GENÉRICA DE ABA
========================================================= */
function getSheet(nomeAba) {
  try {
    if (!nomeAba) {
      throw new Error("Nome da aba não informado.");
    }

    const ss = getSpreadsheet();

    const sheet = ss.getSheetByName(nomeAba);

    if (!sheet) {
      throw new Error(`Aba '${nomeAba}' não encontrada.`);
    }

    return sheet;
  } catch (e) {
    Logger.log(`Erro getSheet(${nomeAba}): ${e.message}`);

    return null;
  }
}

/* =========================================================
   ABA EVENTS
========================================================= */
function getSheetEventos() {
  return getSheet(typeof ABA_EVENTS !== "undefined" ? ABA_EVENTS : "Events");
}

/* =========================================================
   ABA RSVPS
========================================================= */
function getSheetRSVPs() {
  return getSheet(typeof ABA_RSVPS !== "undefined" ? ABA_RSVPS : "RSVPs");
}

/* =========================================================
   ABA CONTRATOS
========================================================= */
function getSheetContracts() {
  return getSheet(
    typeof ABA_CONTRACTS !== "undefined" ? ABA_CONTRACTS : "Contracts",
  );
}

/* =========================================================
   ABA USUÁRIOS
========================================================= */
function getSheetUsuarios() {
  return getSheet(typeof ABA_USERS !== "undefined" ? ABA_USERS : "User");
}
