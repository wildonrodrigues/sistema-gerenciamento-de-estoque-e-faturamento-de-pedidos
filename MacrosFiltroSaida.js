function FormFiltroSaida() {

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

  var Form = HtmlService.createTemplateFromFile("FiltroSaida");

  Form.list = list.map(function(r){
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("FILTRO SAÍDA ESTOQUE").setHeight(600).setWidth(1300);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "FILTRO SAÍDA ESTOQUE");
  
}

function buscaRegistros(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");
  var guiaSaida = planilha.getSheetByName("Saídas");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 3).getValues();

  var ultimaLinha = guiaSaida.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosSaidas = guiaSaida.getRange(2, 1, ultimaLinha, 17).getValues();

  for(var i = 0; i < dadosSaidas.length; i++){

    var Data = Utilities.formatDate(new Date(dadosSaidas[i][1]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosSaidas[i][1] = Data;

  }

  var arrays = {
    dadosProdutos: dadosProdutos,
    dadosSaidas: dadosSaidas,
  }

  return arrays;

}





























