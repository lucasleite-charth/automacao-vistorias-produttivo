// Não execute. 
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Produttivo')
    .addItem('1) Importar relatórios', 'importarRelatorios')
    .addItem('2) Processar vistorias', 'processarVistorias')
    .addItem('3) Preparar e-mails', 'prepararEmailsLembrete')
    .addItem('4) Enviar e-mails pendentes', 'enviarEmailsVistoria')
    .addItem('5) Agendar revisões no calendário', 'criarEventosCalendar')
    .addItem('6) Enviar e-mails com relatorios', 'enviarRelatorioConsolidadoVistorias')
    .addToUi();
}

function executarPipelineCompleto() {
  importarRelatorios();
  processarVistorias();
  prepararEmailsLembrete();
  enviarEmailsVistoria();
  criarEventosCalendar();
  enviarRelatorioConsolidadoVistorias();
}
