function salvarFotoBase64(idConvidado, base64Data, nomeArquivo, idEvento) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    // ==========================================
    // VALIDAÇÕES
    // ==========================================

    if (!idConvidado || !idEvento) {
      throw new Error("Dados obrigatórios ausentes.");
    }

    if (!base64Data) {
      throw new Error("Imagem não enviada.");
    }

    // ==========================================
    // PERMISSÃO
    // ==========================================

    validarPermissaoEvento(idEvento);

    const sheet = getSheetRSVPs();

    if (!sheet) {
      throw new Error("Aba RSVPs não encontrada.");
    }

    const pasta = DriveApp.getFolderById(ID_PASTA_FOTOS);

    const headerMap = createHeaderMap(sheet);

    const dados = sheet.getDataRange().getValues();

    // ==========================================
    // VALIDA TAMANHO
    // ==========================================

    const tamanhoBase64 = base64Data.length;

    const LIMITE = 5 * 1024 * 1024;

    if (tamanhoBase64 > LIMITE) {
      throw new Error("Imagem excede limite permitido.");
    }

    // ==========================================
    // SANITIZA BASE64
    // ==========================================

    const base64Limpa = base64Data.includes(",")
      ? base64Data.split(",")[1]
      : base64Data;

    // ==========================================
    // NOME SEGURO
    // ==========================================

    const nomeSeguro = String(nomeArquivo || "foto")
      .replace(/[^a-zA-Z0-9_\-.]/g, "_")
      .substring(0, 80);

    // ==========================================
    // CRIA BLOB
    // ==========================================

    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Limpa),
      "image/jpeg",
      nomeSeguro,
    );

    // ==========================================
    // SALVA ARQUIVO
    // ==========================================

    const arquivo = pasta.createFile(blob);

    // 🔒 PRIVADO
    arquivo.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);

    const fileId = arquivo.getId();

    const urlFoto = `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;

    // ==========================================
    // PROCURA CONVIDADO
    // ==========================================

    let linha = -1;

    for (let i = 1; i < dados.length; i++) {
      const evt = String(dados[i][headerMap.idevento] || "").trim();

      const guest = String(dados[i][headerMap.idguest] || "").trim();

      if (
        evt === String(idEvento).trim() &&
        guest === String(idConvidado).trim()
      ) {
        linha = i + 1;
        break;
      }
    }

    if (linha === -1) {
      throw new Error("Convidado não localizado.");
    }

    // ==========================================
    // SALVA URL
    // ==========================================

    if (headerMap.photourl !== undefined) {
      sheet.getRange(linha, headerMap.photourl + 1).setValue(urlFoto);
    }

    return {
      success: true,
      mensagem: "Foto salva com sucesso.",
      url: urlFoto,
      fileId: fileId,
    };
  } catch (e) {
    Logger.log("Erro salvarFotoBase64: " + e.message);

    return {
      success: false,
      mensagem: e.message,
    };
  } finally {
    lock.releaseLock();
  }
}
