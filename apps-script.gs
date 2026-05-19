// ─────────────────────────────────────────────
// Google Apps Script — Base de leads da LP
// ─────────────────────────────────────────────
// COMO CONFIGURAR (faça uma vez só):
//
// 1. Acesse: https://sheets.new  → crie uma planilha nova
//    Nomeie as colunas na linha 1:
//    Data | Nome | WhatsApp | E-mail | Segmento | Desafio | Origem
//
// 2. No menu superior: Extensões → Apps Script
//
// 3. Apague o código existente e cole TODO este arquivo
//
// 4. Clique em "Salvar" (ícone de disquete)
//
// 5. Clique em "Implantar" → "Nova implantação"
//    - Tipo: Aplicativo da web
//    - Executar como: Eu (sua conta Google)
//    - Quem pode acessar: Qualquer pessoa
//    → Clique em "Implantar" e autorize quando solicitado
//
// 6. Copie a URL gerada (parece com: https://script.google.com/macros/s/...)
//
// 7. No index.html, substitua o valor de SHEETS_URL por essa URL
//
// 8. Faça git add + commit + push → Vercel atualiza automaticamente
// ─────────────────────────────────────────────

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.nome     || '',
      data.whatsapp || '',
      data.email    || '',
      data.segmento || '',
      data.desafio  || '',
      data.origem   || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Teste manual: rode esta função dentro do Apps Script para verificar se está gravando
function testar() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([new Date(), 'Teste', '(31) 99999-9999', 'teste@email.com', 'E-commerce', 'Teste de integração', 'Script']);
}
