# **Automação de Envio de E-mails para Planilha de Contatos 📧**

Este repositório contém um **script do Google Apps** desenvolvido para automatizar o envio de **e-mails em massa** com base em uma planilha do Google Sheets. O objetivo principal é enviar **50 e-mails por vez**, de forma **automática** e **diária**, a partir de uma lista de contatos.

---

## **🔎 Contexto do Projeto**

Este código foi desenvolvido para a **A.C.E. Consultoria**, com o intuito de automatizar a comunicação com potenciais clientes. A empresa utiliza uma planilha no **Google Sheets** para armazenar dados de contatos de clientes, incluindo nome da empresa, nome do proprietário e e-mail.

### **Objetivo:**
O objetivo principal é automatizar o envio de e-mails para os contatos da planilha, personalizando as mensagens e enviando **50 e-mails por vez**. O código é configurado para ser executado **automaticamente todos os dias**, sem a necessidade de intervenção manual.

> **Link para a planilha**: [Planilha de Contatos](https://docs.google.com/spreadsheets/d/1rXnH9hjqiBF7QGFoU60np5DfPFH5ze9T0_VAYyjQsR0/edit?gid=0#gid=0)

---

## **⚙️ Como Funciona o Código**

O código é escrito em **JavaScript** e utiliza a plataforma **Google Apps Script** para integração com o **Google Sheets** e o **Gmail**. Abaixo está uma explicação detalhada de como o código opera:

### **1. Acessando a Planilha:**
O código começa acessando a planilha ativa no Google Sheets e buscando os dados da aba chamada "Página1". Ele coleta todos os dados da planilha utilizando o método `getDataRange().getValues()`.

```
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");
var rows = sheet.getDataRange().getValues();
```

2. Definindo o Modelo de E-mail:
Um modelo de e-mail é utilizado, e os campos {{Nome}} e {{Empresa}} são substituídos pelos dados da planilha.

```
var emailTemplate = "Olá! {{Nome}}, Tudo bem? Meu nome é Ricardo...";
```

3. Iterando pelas Linhas da Planilha:
O código percorre as linhas da planilha e verifica se o status da linha é diferente de "Enviado". Se não estiver marcado como "Enviado", o e-mail será enviado para aquele contato.

```
for (var i = 1; i < rows.length && count < 50; i++) {
  var status = rows[i][3]; // Coluna 4: Status Mensagem
  if (status === "Enviado") continue;
```

4. Enviando o E-mail:
Se o e-mail for válido, o script envia o e-mail utilizando o MailApp.sendEmail() e marca o status da mensagem como "Enviado" na planilha.

```
MailApp.sendEmail({
  to: email,
  subject: "Proposta da A.C.E. Consultoria",
  body: message,
  bcc: "ricardosilveira@aceconsultoria.com.br"
});
```

5. Atualizando a Planilha:
Após o envio do e-mail, o status da coluna "Status Mensagem" é atualizado para "Enviado", garantindo que o mesmo e-mail não seja enviado novamente.

```
sheet.getRange(i + 1, 4).setValue("Enviado");
```

📅 Automatização Diária
O script é configurado para ser executado automaticamente todos os dias, disparando 50 e-mails por vez. Para isso, é utilizado um gatilho baseado em tempo (Time-driven Trigger), que faz com que o código seja executado em horários definidos sem precisar de ação manual.

Como Configurar o Gatilho:

 - Acesse o editor de Apps Script: Extensões > Apps Script.

 - Clique no ícone de ⏰ Gatilhos na lateral esquerda.

 - Clique em "+ Adicionar Gatilho".

 - Selecione:

    - Função: sendEmails

    - Evento: Baseado em tempo

    - Frequência: Diariamente

    - Hora do dia: Escolha o horário desejado (ex: entre 9h e 10h)

🔑 Considerações e Funcionalidades
1. Validação de E-mails:
O código valida se os e-mails têm o formato correto antes de enviá-los, utilizando uma expressão regular para garantir que o e-mail é válido.

```
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
```

2. Controle de Envio:
O código envia 50 e-mails por vez e marca os e-mails como “Enviado” para garantir que não haja duplicidade no envio. Apenas as linhas não enviadas recebem o e-mail.

3. Personalização das Mensagens:
A mensagem enviada para cada destinatário é personalizada com o nome da pessoa e o nome da empresa, adaptando-se automaticamente a cada contato.

4. Monitoramento de Envio:
O código envia um BCC para um endereço de e-mail de monitoramento (por exemplo, para o responsável do processo), garantindo que a comunicação possa ser monitorada.

📝 Como Configurar o Script
1. Adicionar o Código no Google Apps Script:

 - Abra sua planilha do Google Sheets.

 - Acesse Extensões > Apps Script.

- Apague qualquer código antigo e cole o código fornecido acima.

 - Salve o projeto com um nome, como Automação de Envio de E-mails.

2. Configurar o Gatilho de Execução Automática:
No editor de script, clique no ícone ⏰ para acessar os Gatilhos.

Clique em “+ Adicionar Gatilho” e configure o gatilho para rodar diariamente no horário desejado.

3. Permissões de Execução:
Na primeira execução, o Google pedirá permissões para acessar o Gmail e a Planilha. Permita as permissões solicitadas.

📊 Possíveis Melhorias
Logs de Envio: Registrar a data/hora de cada envio em uma aba separada.

E-mails de Relatório: Enviar um resumo diário para o responsável sobre os e-mails enviados.

Limite de Envio: Implementar uma verificação para garantir que o número máximo de e-mails (100 para contas gratuitas) não seja ultrapassado.
