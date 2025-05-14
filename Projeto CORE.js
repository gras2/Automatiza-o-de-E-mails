function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");
  var rows = sheet.getDataRange().getValues();
  var emailTemplate = "Olá! {{Nome}}, Tudo bem? Meu nome é Ricardo, sou consultor de projetos da A.C.E. Consultoria, empresa de consultoria em gestão empresarial com mais de 30 anos de experiência ajudando empresas a aprimorar sua gestão e superar desafios diversos em diferentes segmentos de mercado. Ao longo dessa trajetória, temos acompanhado de perto a evolução do setor logístico e, com isso, observamos diversas oportunidades de desenvolvimento para empresas como a {{Empresa}}, tendo, inclusive, a oportunidade de alavancar os resultados de diversos negócios do ramo. Com nossa experiência, observamos que muitas empresas desse segmento enfrentam desafios semelhantes, como otimizar o planejamento operacional, melhorar a gestão de processos e documentos, e até mesmo atrair novos clientes de maneira mais eficiente. Acredito que esses pontos podem ser relevantes para vocês, e gostaria de entender melhor como vocês têm lidado com esses desafios. Se for do seu interesse, seria um prazer agendar uma conversa para compartilharmos mais sobre nossa experiência e explorarmos como podemos contribuir para a evolução do seu negócio. Fico à disposição para agendarmos um horário que seja conveniente para você. Agradeço pela atenção e espero poder conversar em breve. Atenciosamente, Gleybson Ricardo, Consultor de projetos | A.C.E. Consultoria";

  var count = 0; // Contador de e-mails enviados

  for (var i = 1; i < rows.length && count < 50; i++) {
    var status = rows[i][3]; // Coluna 4: "Status Mensagem"
    var name = rows[i][0];   // Coluna 1: "Nome cliente"
    var company = rows[i][1]; // Coluna 2: "Empresa"
    var email = rows[i][2];   // Coluna 3: "Email"

    // Pular se já estiver "Enviado"
    if (status === "Enviado") continue;

    email = email ? email.trim().replace(/\s+/g, '') : "";

    if (email && isValidEmail(email)) {
      var message = emailTemplate.replace("{{Nome}}", name)
                                 .replace("{{Empresa}}", company);

      MailApp.sendEmail({
        to: email,
        subject: "Proposta da A.C.E. Consultoria",
        body: message,
        bcc: "ricardosilveira@aceconsultoria.com.br",
      });

      sheet.getRange(i + 1, 4).setValue("Enviado");
      count++;
    } else {
      Logger.log("E-mail inválido na linha " + (i + 1) + ": " + email);
    }
  }
}

// Validação básica de e-mail
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}