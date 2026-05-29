function salvarContratoSaaS(dados) {
  // 1. REMOVIDO: const user = getUsuarioLogado();
  // O GitHub não consegue passar a sessão do Google.

  const sheet =
    SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName("Contracts");

  if (!sheet) {
    throw new Error("Aba Contracts não encontrada.");
  }

  // 2. Usamos o email enviado nos 'dados' ou um valor padrão
  const emailAutor = dados.email || "sistema@api.com";

  sheet.appendRow([
    Utilities.getUuid(),
    dados.cerimonialistaId,
    dados.plano,
    Number(dados.valor) || 0,
    dados.expiracao,
    new Date(),
    emailAutor, // Salva o e-mail que veio do formulário
  ]);

  return { success: true };
}
