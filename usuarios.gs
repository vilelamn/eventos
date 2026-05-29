function listarFuncionariosTabela(usuarioLogadoEmail, roleLogado) {
  try {
    const sheet = getSheetUsuarios();

    if (!sheet) {
      throw new Error("Aba 'User' não encontrada.");
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return [];
    }

    const headerMap = createHeaderMap(sheet);

    const emailBusca = normalizarEmail(usuarioLogadoEmail);
    const roleBusca = normalizarRole(roleLogado);

    const lista = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const emailReg = normalizarEmail(row[headerMap.email]);

      const chefeReg = normalizarEmail(row[headerMap.chefe_email]);

      const roleReg = normalizarRole(row[headerMap.role]);

      let podeVer = false;

      if (["ADMIN", "ADMINISTRADOR"].includes(roleBusca)) {
        podeVer = true;
      } else if (roleBusca === "CERIMONIALISTA") {
        podeVer = emailReg === emailBusca || chefeReg === emailBusca;
      } else {
        podeVer = emailReg === emailBusca;
      }

      if (podeVer) {
        lista.push({
          id_username: row[headerMap.id_username] || "",
          name: row[headerMap.name] || "Sem Nome",
          email: emailReg,
          phone: row[headerMap.phone] || "",
          role: roleReg || "FUNCIONARIO",
        });
      }
    }

    return lista;
  } catch (e) {
    Logger.log("Erro listarFuncionariosTabela: " + e.message);

    return [];
  }
}
function getUserByEmail(email) {
  try {
    const sheet = getSheetUsuarios();

    if (!sheet) return null;

    const headerMap = createHeaderMap(sheet);

    if (!headerMap.email) {
      throw new Error("Coluna email não encontrada.");
    }

    const data = sheet.getDataRange().getValues();

    const emailBusca = normalizarEmail(email);

    for (let i = 1; i < data.length; i++) {
      const emailLinha = normalizarEmail(data[i][headerMap.email]);

      if (emailLinha === emailBusca) {
        return montarObjetoFuncionario(data[i], headerMap);
      }
    }

    return null;
  } catch (e) {
    Logger.log("Erro getUserByEmail: " + e.message);

    return null;
  }
}
function salvarFuncionario(funcionarioData) {
  try {
    const sheet = getSheetUsuarios();

    if (!sheet) {
      return {
        success: false,
        message: "Aba User não encontrada.",
      };
    }

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    const isNewUser = !funcionarioData.id_username;

    let rowIndex = -1;

    // ======================================
    // VALIDA E-MAIL DUPLICADO
    // ======================================

    const emailNovo = normalizarEmail(funcionarioData.email);

    for (let i = 1; i < data.length; i++) {
      const emailAtual = normalizarEmail(data[i][headerMap.email]);

      const idAtual = String(data[i][headerMap.id_username] || "");

      if (
        emailAtual === emailNovo &&
        idAtual !== String(funcionarioData.id_username || "")
      ) {
        return {
          success: false,
          message: "Já existe usuário com este e-mail.",
        };
      }

      if (!isNewUser && idAtual === String(funcionarioData.id_username)) {
        rowIndex = i + 1;
      }
    }

    const novaLinha = Array(sheet.getLastColumn()).fill("");

    // ======================================
    // NOVO USUÁRIO
    // ======================================

    if (isNewUser) {
      novaLinha[headerMap.id_username] = gerarUUID();

      novaLinha[headerMap.date_creation] = new Date();

      novaLinha[headerMap.username] = emailNovo.split("@")[0];
    } else {
      if (rowIndex === -1) {
        return {
          success: false,
          message: "Usuário não encontrado.",
        };
      }

      const linhaAtual = sheet
        .getRange(rowIndex, 1, 1, sheet.getLastColumn())
        .getValues()[0];

      linhaAtual.forEach((v, i) => {
        novaLinha[i] = v;
      });
    }

    // ======================================
    // PREENCHIMENTO DINÂMICO
    // ======================================

    for (const campo in headerMap) {
      if (funcionarioData[campo] !== undefined) {
        let valor = funcionarioData[campo];

        if (typeof valor === "string") {
          valor = valor.trim();
        }

        if (campo === "email" || campo === "chefe_email") {
          valor = normalizarEmail(valor);
        }

        if (campo === "role") {
          valor = normalizarRole(valor);
        }

        if (campo === "pin") {
          valor = String(valor || "").replace(/\.0$/, "");
        }

        novaLinha[headerMap[campo]] = valor;
      }
    }

    // ======================================
    // SALVA
    // ======================================

    if (isNewUser) {
      sheet.appendRow(novaLinha);
    } else {
      sheet.getRange(rowIndex, 1, 1, novaLinha.length).setValues([novaLinha]);
    }

    return {
      success: true,
      message: isNewUser ? "Usuário cadastrado!" : "Usuário atualizado!",
    };
  } catch (e) {
    Logger.log("Erro salvarFuncionario: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}
function excluirFuncionarioNoServidor(id) {
  try {
    const sheet = getSheetUsuarios();

    if (!sheet) {
      return {
        success: false,
        message: "Aba User não encontrada.",
      };
    }

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const idLinha = String(data[i][headerMap.id_username]);

      if (idLinha === String(id)) {
        sheet.deleteRow(i + 1);

        return {
          success: true,
          message: "Usuário excluído.",
        };
      }
    }

    return {
      success: false,
      message: "Usuário não encontrado.",
    };
  } catch (e) {
    Logger.log("Erro excluirFuncionario: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}
function getFuncionarioById(id) {
  try {
    const sheet = getSheetUsuarios();

    if (!sheet) return null;

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const idLinha = String(data[i][headerMap.id_username]);

      if (idLinha === String(id)) {
        return montarObjetoFuncionario(data[i], headerMap);
      }
    }

    return null;
  } catch (e) {
    Logger.log("Erro getFuncionarioById: " + e.message);

    return null;
  }
}
function getCerimonialistas() {
  try {
    const sheet = getSheetUsuarios();

    if (!sheet) {
      return [];
    }

    const headerMap = createHeaderMap(sheet);

    const data = sheet.getDataRange().getValues();

    const lista = [];

    for (let i = 1; i < data.length; i++) {
      const role = normalizarRole(data[i][headerMap.role]);

      if (role === "CERIMONIALISTA") {
        lista.push({
          id: data[i][headerMap.id_username] || "",
          nome: data[i][headerMap.name] || "",
          email: data[i][headerMap.email] || "",
        });
      }
    }

    return lista;
  } catch (e) {
    Logger.log("Erro getCerimonialistas: " + e.message);

    return [];
  }
}
function salvarUsuarioAnfitriao(payload) {
  try {
    const sheetUser = getSheetUsuarios();
    const headerMapUser = createHeaderMap(sheetUser);
    if (!sheetUser) return false;

    const dataUsers = sheetUser.getDataRange().getValues();

    // --- LÓGICA 1: GERAÇÃO DO E-MAIL ÚNICO ---
    const colIdxEmail = headerMapUser["email"];
    const emailsExistentes = dataUsers.map((row) => row[colIdxEmail]);

    const primeiroNome = payload.hosts.split(" ")[0].toLowerCase().trim();
    let emailBase = primeiroNome + "@anfitricao.com";
    let novoEmail = emailBase;
    let contador = 1;

    while (emailsExistentes.includes(novoEmail)) {
      novoEmail = `${primeiroNome}${contador}@anfitriao.com`;
      contador++;
    }

    // --- LÓGICA 2: GERAÇÃO DO ID ÚNICO ---
    const novoIdUnico = Utilities.getUuid();

    // Dados para a aba USER
    const userData = {
      id_username: novoIdUnico,
      name: payload.hosts,
      email: novoEmail,
      role: "CLIENTE",
      pin: payload.hostpassword,
      date_creation: new Date(),
      chefe_email: payload.cadastrado_por,
      registration: "ATIVO",
    };

    let userRowIndex = sheetUser.getLastRow() + 1;

    // Gravação na aba USER
    for (let campo in userData) {
      if (headerMapUser[campo] !== undefined) {
        sheetUser
          .getRange(userRowIndex, headerMapUser[campo] + 1)
          .setValue(userData[campo]);
      }
    }

    // --- NOVA LÓGICA: GRAVAÇÃO NA ABA EVENTS (COLUNA M) ---
    // 1. Aceder à aba Events
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEvents = ss.getSheetByName("Events");

    if (sheetEvents && payload.idevento) {
      const dataEvents = sheetEvents.getDataRange().getValues();
      const idProcurado = payload.idevento;
      let linhaEvento = -1;

      // 2. Procurar a linha que tem o ID do evento (Coluna A)
      for (let i = 1; i < dataEvents.length; i++) {
        if (dataEvents[i][0] == idProcurado) {
          linhaEvento = i + 1; // +1 porque o array começa em 0 e a planilha em 1
          break;
        }
      }

      // 3. Se encontrar o evento, grava o e-mail na Coluna M (13)
      if (linhaEvento !== -1) {
        sheetEvents.getRange(linhaEvento, 13).setValue(novoEmail);
        console.log(
          "E-mail duplicado na aba Events para o evento: " + idProcurado,
        );
      } else {
        console.warn(
          "Evento não encontrado na aba Events para gravar o e-mail.",
        );
      }
    }

    return true;
  } catch (e) {
    console.error("Erro no salvarUsuarioAnfitriao: " + e.message);
    return false;
  }
}
function getEmailsDoGrupoCerimonial_(userEmail, userRole) {
  const email = String(userEmail || "")
    .toLowerCase()
    .trim();
  const role = String(userRole || "")
    .toUpperCase()
    .trim();

  const sheetUsuarios =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User");
  if (!sheetUsuarios) return [email];

  const dados = sheetUsuarios.getDataRange().getValues();
  const cab = dados[0].map((c) => c.toLowerCase().trim());
  const idxEmail = cab.indexOf("email");
  const idxChefe = cab.indexOf("chefe_email");

  let chefeEmail = email;

  // Se for funcionário ou recepcionista, descobre o chefe
  if (role === "FUNCIONARIO" || role === "RECEPCIONISTA") {
    const linha = dados
      .slice(1)
      .find((r) => String(r[idxEmail]).toLowerCase().trim() === email);
    if (linha && idxChefe !== -1) {
      chefeEmail = String(linha[idxChefe] || email)
        .toLowerCase()
        .trim();
    }
  }

  // Lista emails do grupo
  const grupo = [chefeEmail];

  dados.slice(1).forEach((r) => {
    if (String(r[idxChefe]).toLowerCase().trim() === chefeEmail) {
      grupo.push(String(r[idxEmail]).toLowerCase().trim());
    }
  });

  return [...new Set(grupo)];
}
