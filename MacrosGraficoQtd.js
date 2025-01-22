function FormGraficoQtd() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosLinha = guiaProduto.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLinha.flat())];

  var listaLinhas = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLinhas.push([listaUnica[i]]);
  }

  dadosLinha.length = 0;
  listaUnica.length = 0;

  var list = listaLinhas.sort();

  var Form = HtmlService.createTemplateFromFile("GraficoQtd");

  Form.list = list.map(function(r){
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("GRÁFICO QTD").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "GRÁFICO QTD");
  
}

