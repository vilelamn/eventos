/* =========================================================
   MAPA DE CABEÇALHOS
========================================================= */
function createHeaderMap(sheet) {
  try {
    if (!sheet) {
      throw new Error("Sheet não informada.");
    }

    const lastCol = sheet.getLastColumn();

    if (lastCol <= 0) {
      return {};
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    const map = {};

    headers.forEach((header, index) => {
      if (
        header === null ||
        header === undefined ||
        String(header).trim() === ""
      ) {
        return;
      }

      const normalized = String(header)
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .replace(/[\s\-]+/g, "_")
        .replace(/[^a-z0-9_]/g, "")
        .replace(/^_+|_+$/g, "");

      if (!normalized) {
        return;
      }

      if (map[normalized] !== undefined) {
        Logger.log(`⚠️ Cabeçalho duplicado: "${normalized}"`);

        return;
      }

      map[normalized] = index;
    });

    return map;
  } catch (e) {
    Logger.log("Erro createHeaderMap: " + e.message);

    return {};
  }
}

/* =========================================================
   NORMALIZADORES
========================================================= */
function normalizarEmail(email) {
  return String(email || "")
    .trim()
    .toLowerCase();
}

function normalizarRole(role) {
  return String(role || "")
    .trim()
    .toUpperCase();
}

function normalizarTexto(texto) {
  return String(texto || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

/* =========================================================
   UUID
========================================================= */
function gerarUUID() {
  return Utilities.getUuid();
}

/* =========================================================
   FORMATAR DATA
========================================================= */
function formatarData(data, formato = "dd/MM/yyyy") {
  try {
    if (!(data instanceof Date) || isNaN(data.getTime())) {
      return "";
    }

    const timezone =
      typeof TIMEZONE !== "undefined" ? TIMEZONE : Session.getScriptTimeZone();

    return Utilities.formatDate(data, timezone, formato);
  } catch (e) {
    Logger.log("Erro formatarData: " + e.message);

    return "";
  }
}

/* =========================================================
   VALIDAR EMAIL
========================================================= */
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

/* =========================================================
   INCLUDE HTML
========================================================= */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    Logger.log(`Erro include(${filename}): ${e.message}`);

    return "";
  }
}

/* =========================================================
   MONTAR OBJETO DINÂMICO
========================================================= */
function montarObjetoFuncionario(row, headerMap) {
  try {
    if (!row || !headerMap) {
      return {};
    }

    const obj = {};

    for (const key in headerMap) {
      let value = row[headerMap[key]];

      /* =========================================
         DATAS
      ========================================= */
      if (value instanceof Date) {
        value = Utilities.formatDate(
          value,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm:ss",
        );
      } else if (typeof value === "number" && Number.isInteger(value)) {

      /* =========================================
         NORMALIZA NÚMEROS GRANDES
      ========================================= */
        value = String(value);
      }

      obj[key] = value !== null && value !== undefined ? value : "";
    }

    /* =========================================
       NORMALIZAÇÕES
    ========================================= */
    if (obj.role) {
      obj.role = normalizarRole(obj.role);
    }

    if (obj.email) {
      obj.email = normalizarEmail(obj.email);
    }

    if (obj.chefe_email) {
      obj.chefe_email = normalizarEmail(obj.chefe_email);
    }

    return obj;
  } catch (e) {
    Logger.log("Erro montarObjetoFuncionario: " + e.message);

    return {};
  }
}

/* =========================================================
   CONFIGURAÇÕES GLOBAIS
========================================================= */
function getGlobalConfigs() {
  return {
    apiUrl: ScriptApp.getService().getUrl(),

    versao: "1.0.3",

    timezone:
      typeof TIMEZONE !== "undefined" ? TIMEZONE : Session.getScriptTimeZone(),
  };
}

/* =========================================================
   URL WEBAPP
========================================================= */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/* =========================================================
   SPREADSHEET
========================================================= */
function getSS() {
  if (!ID_DA_PLANILHA) {
    throw new Error("ID_DA_PLANILHA não definido.");
  }

  return SpreadsheetApp.openById(ID_DA_PLANILHA);
}

/* =========================================================
   GET SHEET SAFE
========================================================= */
function getSheetByNameSafe(name) {
  try {
    if (!name) {
      throw new Error("Nome da aba não informado.");
    }

    const sheet = getSS().getSheetByName(name);

    if (!sheet) {
      throw new Error(`Aba '${name}' não encontrada.`);
    }

    return sheet;
  } catch (e) {
    Logger.log("Erro getSheetByNameSafe: " + e.message);

    return null;
  }
}
