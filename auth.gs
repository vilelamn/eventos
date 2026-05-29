/* =========================================================
   AUTENTICAÇÃO DE ACESSO
========================================================= */

function autenticarAcesso(email, pin) {
  try {
    if (!email || !pin) {
      return {
        success: false,
        reason: "DADOS_INVALIDOS",
      };
    }

    const emailNormalizado = String(email).toLowerCase().trim();

    const pinDigitado = String(pin).trim();

    const user = getUserByEmail(emailNormalizado);

    if (!user) {
      return {
        success: false,
        reason: "EMAIL_NOT_FOUND",
      };
    }

    /* =====================================================
       HASH DO PIN DIGITADO
    ===================================================== */

    const pinHash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pinDigitado),
    );

    const pinPlanilha = String(user.pin || "").trim();

    /* =====================================================
       COMPARAÇÃO SEGURA
    ===================================================== */

    if (pinPlanilha !== pinHash) {
      return {
        success: false,
        reason: "PIN_INVALID",
      };
    }

    /* =====================================================
       USUÁRIO BLOQUEADO
    ===================================================== */

    const registro = String(user.registration || "ATIVO")
      .toUpperCase()
      .trim();

    if (registro === "BLOQUEADO") {
      return {
        success: false,
        reason: "USER_BLOCKED",
      };
    }

    return {
      success: true,

      email: user.email || "",

      role: user.role || "VISITANTE",

      name: user.name || "",

      chefeEmail: user.chefe_email || "",
    };
  } catch (e) {
    Logger.log("Erro autenticarAcesso: " + e.message);

    return {
      success: false,
      message: e.toString(),
    };
  }
}

/* =========================================================
   USUÁRIO LOGADO
========================================================= */

function getUsuarioLogado() {
  let activeUserEmail = "";

  try {
    if (Session && Session.getActiveUser) {
      activeUserEmail = Session.getActiveUser().getEmail() || "";
    }
  } catch (e) {
    Logger.log("ERRO Session.getActiveUser(): " + e.message);

    return {
      success: false,
      email: "",
      role: "VISITANTE",
      message: "Falha ao obter sessão.",
    };
  }

  activeUserEmail = String(activeUserEmail).trim().toLowerCase();

  if (!activeUserEmail) {
    Logger.log("Usuário sem sessão ativa.");

    return {
      success: false,
      email: "",
      role: "VISITANTE",
      message: "Usuário não autenticado.",
    };
  }

  const userRecord = getUserByEmail(activeUserEmail);

  if (!userRecord) {
    Logger.log("Usuário não encontrado: " + activeUserEmail);

    return {
      success: false,
      email: activeUserEmail,
      role: "VISITANTE",
      message: "Usuário não cadastrado.",
    };
  }

  /* =====================================================
     USUÁRIO BLOQUEADO
  ===================================================== */

  const registro = String(userRecord.registration || "ATIVO")
    .toUpperCase()
    .trim();

  if (registro === "BLOQUEADO") {
    return {
      success: false,
      email: activeUserEmail,
      role: "VISITANTE",
      message: "Usuário bloqueado.",
    };
  }

  Logger.log("Usuário autenticado: " + activeUserEmail);

  return {
    success: true,

    id_username: userRecord.id_username || userRecord.id_user || "",

    name: userRecord.name || "",

    username: userRecord.username || "",

    email: userRecord.email || "",

    role: userRecord.role || "VISITANTE",

    chefeEmail: userRecord.chefe_email || "",
  };
}

/* =========================================================
   ADMIN
========================================================= */

function isUserAdmin() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();

    return email === String(EMAIL_ADMIN).toLowerCase().trim();
  } catch (e) {
    Logger.log("Erro isUserAdmin: " + e.message);

    return false;
  }
}

/* =========================================================
   RECUPERAÇÃO DE PIN
========================================================= */

function enviarRecuperacaoSenha(email) {
  try {
    if (!email) {
      return "E-mail inválido.";
    }

    const emailLower = String(email).trim().toLowerCase();

    const sheet = getSheetUsuarios();

    if (!sheet) {
      return "Planilha de usuários não encontrada.";
    }

    if (sheet.getLastRow() <= 1) {
      return "Nenhum usuário cadastrado.";
    }

    const headerMap = createHeaderMap(sheet);

    if (
      headerMap.email === undefined ||
      headerMap.pin === undefined ||
      headerMap.name === undefined
    ) {
      Logger.log("Colunas obrigatórias não encontradas.");

      return "Erro interno no sistema.";
    }

    const dados = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .getValues();

    let linhaUsuario = -1;
    let usuario = null;

    for (let i = 0; i < dados.length; i++) {
      const emailPlanilha = String(dados[i][headerMap.email] || "")
        .trim()
        .toLowerCase();

      if (emailPlanilha === emailLower) {
        linhaUsuario = i + 2;
        usuario = dados[i];

        break;
      }
    }

    // 🔐 Não revela existência do usuário
    if (!usuario) {
      return "Se o e-mail existir, as instruções foram enviadas.";
    }

    /* =====================================================
       GERA NOVO PIN
    ===================================================== */

    const novoPin = Math.floor(1000 + Math.random() * 9000).toString();

    /* =====================================================
       HASH DO NOVO PIN
    ===================================================== */

    const novoPinHash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, novoPin),
    );

    sheet.getRange(linhaUsuario, headerMap.pin + 1).setValue(novoPinHash);

    const nomeUsuario = usuario[headerMap.name]
      ? String(usuario[headerMap.name]).split(" ")[0]
      : "Olá";

    MailApp.sendEmail({
      to: emailLower,

      subject: "Recuperação de PIN - MeuEvento",

      htmlBody: `
        <div style="font-family:Arial;padding:20px;">
          <h2>Recuperação de PIN</h2>

          <p>Olá ${nomeUsuario},</p>

          <p>Seu novo PIN de acesso é:</p>

          <div style="
            font-size:28px;
            font-weight:bold;
            color:#4f46e5;
            padding:15px;
            border:1px solid #ddd;
            display:inline-block;
            border-radius:8px;
            margin:10px 0;
          ">
            ${novoPin}
          </div>

          <p>
            Utilize este PIN para acessar o sistema.
          </p>

          <p>
            Caso não tenha solicitado,
            ignore este e-mail.
          </p>
        </div>
      `,
    });

    return "PIN redefinido com sucesso.";
  } catch (e) {
    Logger.log("Erro enviarRecuperacaoSenha: " + e.message);

    return "Erro interno ao recuperar senha.";
  }
}
