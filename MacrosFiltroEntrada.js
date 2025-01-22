function FormFiltroEntrada() {

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

  var Form = HtmlService.createTemplateFromFile("FiltroEntrada");

  Form.list = list.map(function(r){
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("FILTRO ENTRADA ESTOQUE").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "FILTRO ENTRADA ESTOQUE");
  
}

function buscaDados(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");
  var guiaEntrada = planilha.getSheetByName("Entradas");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 3).getValues();

  var ultimaLinha = guiaEntrada.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosEntradas = guiaEntrada.getRange(2, 1, ultimaLinha, 16).getValues();

  for(var i = 0; i < dadosEntradas.length; i++){

    var Data = Utilities.formatDate(new Date(dadosEntradas[i][1]), planilha.getSpreadsheetTimeZone(), "dd/MM/yyyy");

    dadosEntradas[i][1] = Data;

  }

  var arrays = {
    dadosProdutos: dadosProdutos,
    dadosEntradas: dadosEntradas,
  }

  return arrays;

}


















