function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu('Formulários')
  .addItem('Cadastro Produto', 'FormProduto')
  .addItem('Entrada', 'FormEntrada')
  .addItem('Saída', 'FormSaida')
  .addItem('Filtro Entrada', 'FormFiltroEntrada')
  .addItem('Filtro Saída', 'FormFiltroSaida')
  .addItem('Relatório', 'FormRelatorio')
  .addItem('Gráfico Qtd', 'FormGraficoQtd')
  .addItem('Gráfico Valor', 'FormGraficoValor')
  .addToUi();
}

function FormProduto() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var dadosLinha = guiaProduto.getRange(2,1,ultimaLinha,1).getValues();

  var b = {};

  for (var i = 0; i < dadosLinha.length; i++){
    b[dadosLinha[i][0]] = dadosLinha[i][0];
  }

  var listaLinhas = [];

  for(var key in b){
    listaLinhas.push([key]);
  }

  dadosLinha.length = 0;

  var list = listaLinhas.sort();
  
  var Form = HtmlService.createTemplateFromFile("FormProduto");

  Form.list = list.map(function(r){
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("CADASTRO DE PRODUTOS").setHeight(320).setWidth(390);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"CADASTRO LOGÍSTICO");

}


function Chamar(Arquivo){
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}

function SalvarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Linha = Dados.Linha;
    var Marca = Dados.Marca;
    var Produto = Dados.Produto;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Produtos");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,3).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Marca && dadosProdutos[i][2] == Produto){

        dadosProdutos.length = 0;
        return "JÁ CADASTRADO!";

      }

    }

    var linha = ultimaLinha + 1;

    guiaProduto.getRange(linha,1).setValue(Linha);
    guiaProduto.getRange(linha,2).setValue(Marca);
    guiaProduto.getRange(linha,3).setValue(Produto);

    guiaProduto.getRange("A2:C").sort([{column: 1, ascending: true},{column: 2, ascending: true},{column: 3, ascending: true}]);

   dadosProdutos.length = 0;
   return "REGISTRADO COM SUCESSO!";

  }

}


function AtualizarListaLinhas(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var dadosLinha = guiaProduto.getRange(2,1,ultimaLinha,1).getValues();

  var b = {};

  for (var i = 0; i < dadosLinha.length; i++){
    b[dadosLinha[i][0]] = dadosLinha[i][0];
  }

  var listaLinhas = [];

  for(var key in b){
    listaLinhas.push([key]);
  }

  dadosLinha.length = 0;

  var list = listaLinhas.sort();

  return list;

}


function listaMarcas(Linha){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosMarcas = guiaProduto.getRange(2,1,ultimaLinha,2).getValues();

  var marcas = [];

  for(var i = 0; i <dadosMarcas.length; i++){

    if(dadosMarcas[i][0] == Linha){
      var p = dadosMarcas[i][1];
      marcas.push(p);
    }

  }

  var b = {};

  for(var i = 0; i < marcas.length; i++){
    b[marcas[i]] = marcas[i];
  }

  var listaMarcas = [];

  for(var key in b){
    listaMarcas.push([key]);
  }

  dadosMarcas.length = 0;
  marcas.length = 0;

  return listaMarcas.sort();

}


function listaProdutos(Linha,Marca){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2, 1,ultimaLinha,3).getValues();

  var produtos = [];

  for(var i = 0; i < dadosProdutos.length; i++){

    if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Marca){
      produtos.push([dadosProdutos[i][2]]);
    }

  }

  dadosProdutos.length = 0;

  return produtos.sort();
     
}


function EditarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var LinhaLista = Dados.LinhaLista;
    var MarcaLista = Dados.MarcaLista;
    var ProdutoLista = Dados.ProdutoLista;    
    var Linha = Dados.Linha;
    var Marca = Dados.Marca;
    var Produto = Dados.Produto;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Produtos");
    var guiaEntrada = planilha.getSheetByName("Entradas");

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntradas = guiaEntrada.getRange(2, 5, ultimaLinha, 3).getValues();

    var ver = dadosEntradas.filter(function(value,i,arr){
      return LinhaLista == arr[i][0] && MarcaLista == arr[i][1] && ProdutoLista == arr[i][2];
    });

    if(ver.length > 0){
      dadosEntradas.length = 0;
      ver.length = 0;         
      return "NÃO PODE SER EDITADO. JÁ POSSUI LANÇAMENTO DE ENTRADA.";
    }

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 3).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == LinhaLista && dadosProdutos[i][1] == MarcaLista && dadosProdutos[i][2] == ProdutoLista){

        var linha = i + 2;

        guiaProduto.getRange(linha,1).setValue(Linha);
        guiaProduto.getRange(linha,2).setValue(Marca);
        guiaProduto.getRange(linha,3).setValue(Produto);

        dadosEntradas.length = 0;
        dadosProdutos.length = 0;
        ver.length = 0;

        return "EDITADO COM SUCESSO!";

      }

    }

    dadosEntradas.length = 0;
    dadosProdutos.length = 0; 
    ver.length = 0;   
    return "NÃO ENCONTRADO!";

  }

}

function ExcluirProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Linha = Dados.Linha;
    var Marca = Dados.Marca;
    var Produto = Dados.Produto;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Produtos");
    var guiaEntrada = planilha.getSheetByName("Entradas");

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntradas = guiaEntrada.getRange(2, 5, ultimaLinha, 3).getValues();

    var ver = dadosEntradas.filter(function(value, i, arr){
      return Linha == arr[i][0] && Marca == arr[i][1] && Produto == arr[i][2];
    });

    if(ver.length > 0){
      dadosEntradas.length = 0;
      ver.length = 0;
      return "NÃO PODE SER EXCLUÍDO. JÁ TEM LANÇAMENTO DE ENTRADA!";
    }

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,3).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Marca && dadosProdutos[i][2] == Produto){

        var linha = i + 2;
        guiaProduto.deleteRow(linha);

        dadosEntradas.length = 0;
        dadosProdutos.length = 0;
        ver.length = 0;
        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dadosEntradas.length = 0;
    dadosProdutos.length = 0;
    ver.length = 0;

    return "PRODUTO NÃO ENCONTRADO!";

  }

}

