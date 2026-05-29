/* =========================================================
   EMAIL.GS
   ENVIO DE E-MAILS DO SISTEMA
========================================================= */

/* =========================================================
   TEMPLATE BASE
========================================================= */

function getTemplateEmail(titulo, conteudo) {
  return `
    <div style="
      font-family: Arial, sans-serif;
      background:#f4f4f5;
      padding:30px;
    ">
      <div style="
        max-width:650px;
        margin:auto;
        background:#ffffff;
        border-radius:12px;
        overflow:hidden;
        box-shadow:0 2px 10px rgba(0,0,0,.08);
      ">
        
        <div style="
          background:#4f46e5;
          padding:25px;
          text-align:center;
        ">
          <h1 style="
            color:#fff;
            margin:0;
            font-size:24px;
          ">
            ${titulo}
          </h1>
        </div>

        <div style="padding:30px;">
          ${conteudo}
        </div>

        <div style="
          padding:20px;
          text-align:center;
          background:#fafafa;
          color:#666;
          font-size:12px;
          border-top:1px solid #eee;
        ">
          MeuEvento © ${new Date().getFullYear()}
        </div>

      </div>
    </div>
  `;
}

/* =========================================================
   ENVIO SEGURO
========================================================= */

function enviarEmail(destinatario, assunto, htmlBody) {
  try {
    if (!destinatario) {
      throw new Error("Destinatário não informado.");
    }

    MailApp.sendEmail({
      to: destinatario,
      subject: assunto,
      htmlBody: htmlBody,
      name: "MeuEvento",
    });

    return {
      success: true,
    };
  } catch (e) {
    Logger.log("Erro enviarEmail: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

/* =========================================================
   E-MAIL CRIAÇÃO EVENTO
========================================================= */

function enviarEmailCriacaoEvento(payload, urlConfirmacao) {
  try {
    let dataFormatada = payload.eventdate || "";

    if (dataFormatada.includes("-")) {
      const partes = dataFormatada.split("-");
      dataFormatada = `${partes[2]}/${partes[1]}/${partes[0]}`;
    }

    const senhaHost = payload.hostpassword || "Não definida";

    const conteudo = `
      <p style="font-size:16px;color:#333;">
        Olá,
      </p>

      <p style="font-size:15px;color:#555;">
        O evento <strong>${payload.eventyname}</strong> foi criado com sucesso.
      </p>

      <div style="
        background:#f8fafc;
        border:1px solid #e5e7eb;
        border-radius:10px;
        padding:20px;
        margin-top:20px;
      ">
        <p><strong>Evento:</strong> ${payload.eventyname}</p>
        <p><strong>Data:</strong> ${dataFormatada}</p>
        <p><strong>Hora:</strong> ${payload.eventhour || ""}</p>
        <p><strong>Local:</strong> ${payload.eventlocation || ""}</p>
        <p><strong>Anfitrião:</strong> ${payload.hosts || ""}</p>
      </div>

      <div style="
        background:#eef2ff;
        border-left:4px solid #4f46e5;
        padding:20px;
        margin-top:25px;
        border-radius:8px;
      ">
        <p style="margin:0 0 10px 0;">
          <strong>Senha do anfitrião:</strong>
        </p>

        <div style="
          font-size:22px;
          font-weight:bold;
          color:#4338ca;
          letter-spacing:2px;
        ">
          ${senhaHost}
        </div>
      </div>

      <div style="
        margin-top:25px;
        padding:20px;
        background:#f0fdf4;
        border:1px dashed #22c55e;
        border-radius:10px;
      ">
        <p>
          <strong>Link RSVP:</strong>
        </p>

        <a href="${urlConfirmacao}" target="_blank">
          ${urlConfirmacao}
        </a>
      </div>
    `;

    return enviarEmail(
      payload.cadastrado_por,
      `Novo Evento: ${payload.eventyname}`,
      getTemplateEmail("Evento Criado", conteudo),
    );
  } catch (e) {
    Logger.log("Erro enviarEmailCriacaoEvento: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

/* =========================================================
   AGRADECIMENTO RSVP
========================================================= */

function enviarEmailAgradecimento(nomeConvidado, emailConvidado, qrCodeUrl) {
  try {
    if (!emailConvidado) {
      return {
        success: false,
        message: "E-mail do convidado não informado.",
      };
    }

    let idConvidado = "";

    try {
      idConvidado = decodeURIComponent(
        qrCodeUrl.split("text=")[1].split("&")[0],
      );
    } catch (_) {}

    const conteudo = `
      <p style="font-size:16px;">
        Olá <strong>${nomeConvidado}</strong>,
      </p>

      <p>
        Sua presença foi confirmada com sucesso.
      </p>

      <div style="
        text-align:center;
        margin:30px 0;
      ">
        <p>
          <strong>Apresente este QR Code na entrada:</strong>
        </p>

        <img 
          src="${qrCodeUrl}"
          width="220"
          height="220"
          style="
            border:8px solid #fff;
            box-shadow:0 2px 10px rgba(0,0,0,.15);
            border-radius:10px;
          "
        />

        <p style="
          margin-top:15px;
          font-size:14px;
          color:#666;
        ">
          ID: ${idConvidado}
        </p>
      </div>

      <div style="
        background:#fafafa;
        border-radius:10px;
        padding:20px;
      ">
        <p><strong>Importante:</strong></p>

        <ul style="line-height:1.8;">
          <li>Chegue com antecedência.</li>
          <li>Tenha este QR Code em mãos.</li>
          <li>Apresente na recepção.</li>
        </ul>
      </div>
    `;

    return enviarEmail(
      emailConvidado,
      "Confirmação de Presença",
      getTemplateEmail("Presença Confirmada 🎉", conteudo),
    );
  } catch (e) {
    Logger.log("Erro enviarEmailAgradecimento: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

/* =========================================================
   ENVIO CONVITES QR CODE
========================================================= */

function enviarConvitesQrCode(emailAnfitriao) {
  try {
    const ss = SpreadsheetApp.openById(ID_DA_PLANILHA);

    const abaEventos = ss.getSheetByName(ABA_EVENTS);
    const abaRSVPs = ss.getSheetByName(ABA_RSVPS);

    if (!abaEventos || !abaRSVPs) {
      throw new Error("Abas obrigatórias não encontradas.");
    }

    const mapEventos = createHeaderMap(abaEventos);
    const mapRSVPs = createHeaderMap(abaRSVPs);

    const dadosEventos = abaEventos.getDataRange().getValues();

    let idEvento = null;

    for (let i = 1; i < dadosEventos.length; i++) {
      const emailHost = normalizarEmail(
        dadosEventos[i][mapEventos.emailanfitriao],
      );

      if (emailHost === normalizarEmail(emailAnfitriao)) {
        idEvento = dadosEventos[i][mapEventos.idevento];
        break;
      }
    }

    if (!idEvento) {
      throw new Error("Evento do anfitrião não encontrado.");
    }

    const dadosRSVPs = abaRSVPs.getDataRange().getValues();

    let enviados = 0;

    for (let i = 1; i < dadosRSVPs.length; i++) {
      const row = dadosRSVPs[i];

      const evt = String(row[mapRSVPs.idevento] || "").trim();

      if (evt !== String(idEvento).trim()) {
        continue;
      }

      const status = String(row[mapRSVPs.status] || "")
        .toUpperCase()
        .trim();

      if (status !== "CONFIRMADO") {
        continue;
      }

      const email = String(row[mapRSVPs.guestemail] || "").trim();

      const nome = String(row[mapRSVPs.guestname] || "").trim();

      const qrCode = String(row[mapRSVPs.qrcode] || "").trim();

      if (!email || !qrCode) {
        continue;
      }

      const conteudo = `
        <p>
          Olá <strong>${nome}</strong>,
        </p>

        <p>
          Seu acesso ao evento está confirmado.
        </p>

        <div style="text-align:center;margin:30px 0;">
          <img 
            src="${qrCode}"
            width="240"
            style="
              border-radius:10px;
              box-shadow:0 2px 10px rgba(0,0,0,.15);
            "
          />
        </div>

        <p style="
          text-align:center;
          color:#666;
        ">
          Apresente este QR Code na entrada.
        </p>
      `;

      enviarEmail(
        email,
        "Seu Convite com QR Code",
        getTemplateEmail("Convite Confirmado", conteudo),
      );

      enviados++;
    }

    return {
      success: true,
      enviados,
      message: `Convites enviados: ${enviados}`,
    };
  } catch (e) {
    Logger.log("Erro enviarConvitesQrCode: " + e.message);

    return {
      success: false,
      message: e.message,
    };
  }
}

/* =========================================================
   RECUPERAÇÃO DE SENHA
========================================================= */

function enviarRecuperacaoSenha(email) {
  try {
    const usuario = buscarUsuarioPorEmail(email);

    if (!usuario) {
      throw new Error("Usuário não encontrado.");
    }

    const conteudo = `
      <p>
        Olá <strong>${usuario.nome || "usuário"}</strong>,
      </p>

      <p>
        Seus dados de acesso:
      </p>

      <div style="
        background:#f8fafc;
        padding:20px;
        border-radius:10px;
        border:1px solid #e5e7eb;
      ">
        <p><strong>E-mail:</strong> ${usuario.email}</p>
        <p><strong>Senha:</strong> ${usuario.pin}</p>
      </div>

      <p style="
        margin-top:25px;
        color:#666;
        font-size:13px;
      ">
        Recomendamos alterar sua senha após o login.
      </p>
    `;

    enviarEmail(
      usuario.email,
      "Recuperação de Acesso",
      getTemplateEmail("Recuperação de Senha", conteudo),
    );

    return "E-mail enviado com sucesso.";
  } catch (e) {
    Logger.log("Erro enviarRecuperacaoSenha: " + e.message);

    return e.message;
  }
}
