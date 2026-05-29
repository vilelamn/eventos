```javascript
// ==========================================
// CONFIGURAÇÕES GLOBAIS
// ==========================================

const CONFIG = {
  PLANILHA_ID: "11DMrzOzOSWlYRrRnLku5TnoaobAQfGfLOjDDr3e6_GY";
  PASTA_FOTOS_ID:"1uHQS3GLHGFWUXn-eCj1s2DTb-kpFAxCl";

  EMAIL_ADMIN: "vilelamn@gmail.com",

  TIMEZONE: Session.getScriptTimeZone(),

  ABAS: {
    USUARIOS: "User",
    EVENTOS: "Events",
    RSVPS: "RSVPs",
    CONTRATOS: "Contracts",
  },

  ROLES: {
    ADMIN: "ADMIN",
    ADMINISTRADOR: "ADMINISTRADOR",
    CERIMONIALISTA: "CERIMONIALISTA",
    HOST: "HOST",
    CLIENTE: "CLIENTE",
  },
};

// ==========================================
// VALIDAÇÃO INICIAL
// ==========================================

function validarConfiguracoes() {
  if (!CONFIG.PLANILHA_ID) {
    throw new Error("PLANILHA_ID não configurado.");
  }

  if (!CONFIG.PASTA_FOTOS_ID) {
    throw new Error("PASTA_FOTOS_ID não configurado.");
  }
}

// ==========================================
// HELPERS GLOBAIS
// ==========================================

function getSpreadsheet() {
  validarConfiguracoes();

  return SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
}

function logInfo(msg) {
  Logger.log("INFO: " + msg);
}

function logErro(msg) {
  Logger.log("ERRO: " + msg);
}
```;
