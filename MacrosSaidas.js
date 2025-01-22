function FormSaida(Id) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");
  var ultimaLinha = guiaProduto.getLastRow();

  if (ultimaLinha == 0) {
    ultimaLinha = 1;
  }

  var dadosLinha = guiaProduto.getRange(2, 1, ultimaLinha, 1).getValues();
  var listaUnica = [...new Set(dadosLinha.flat())];
  var listaLinhas = [];

  for (var i = 0; i < listaUnica.length; i++) {
    listaLinhas.push([listaUnica[i]]);
  }

  dadosLinha.length = 0;
  listaUnica.length = 0;

  var list = listaLinhas.sort();
  var Form = HtmlService.createTemplateFromFile("FormSaida");
  Form.list = list.map(function (r) {
    return r[0];
  });

  Form.Id = Id; 
  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  MostrarForm.setTitle("SAÍDA ESTOQUE").setHeight(500).setWidth(730);
  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "SAÍDA ESTOQUE");
}




function Saldo(Dados){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaEntrada = planilha.getSheetByName("Entradas");
  var guiaSaida = planilha.getSheetByName("Saídas");

  var Linha = Dados.Linha;
  var Marca = Dados.Marca;
  var Produto = Dados.Produto;
  var IdAtual = Dados.IdAtual;

  var ultimaLinha = guiaEntrada.getLastRow();
  var dadosEntradas = guiaEntrada.getRange(2,1, ultimaLinha, 16).getValues();

  var Entradas = dadosEntradas.filter(function(value, i, arr){
    return Linha == arr[i][4] && Marca == arr[i][5] && Produto == arr[i][6];
  });

  if(Entradas.length == 0){
    dadosEntradas.length = 0;
    return "NÃO TEM ENTRADA PARA ESTE PRODUTO!";
  }

  var ultimaLinha = guiaSaida.getLastRow();
  var dadosSaidas = guiaSaida.getRange(2,1, ultimaLinha, 17).getValues();

  var Saidas = dadosSaidas.filter(function(value,i,arr){
    return Linha == arr[i][4] && Marca == arr[i][5] && Produto == arr[i][6];
  });

  var QtdEntrada = 0;
  var QtdSaida = 0;

  for(var i = 0; i < Entradas.length; i++){
    if(Entradas[i][10] != ""){
      QtdEntrada = parseFloat(QtdEntrada) + parseFloat(Entradas[i][10]);
    }
  }

  for(var i = 0; i < Saidas.length; i++){
    if(Saidas[i][9] != ""){
      QtdSaida = parseFloat(QtdSaida) + parseFloat(Saidas[i][9]);
    }
  }

  var SaldoTotal = parseFloat(QtdEntrada) - parseFloat(QtdSaida);
  SaldoTotal = SaldoTotal.toLocaleString({style: 'decimal', decimal: 'pt-BR'});
  SaldoTotal = SaldoTotal.replace(/\./g,"");

  for(var i = 0; i < Entradas.length; i++){

    var IdEntrada = Entradas[i][0];

    if(IdEntrada > IdAtual){

    var QtdIdEntrada = Entradas[i][10];
    var SaidaIdEntrada = 0;
    var SaldoIdEntrada = 0;

    for(var l = 0; l < Saidas.length; l++){
      if(Saidas[l][7] == IdEntrada){
        SaidaIdEntrada = parseFloat(SaidaIdEntrada) + parseFloat(Saidas[l][9]);
      }
    }

    if(parseFloat(QtdIdEntrada) > parseFloat(SaidaIdEntrada)){
      SaldoIdEntrada = parseFloat(QtdIdEntrada) - parseFloat(SaidaIdEntrada);
      SaldoIdEntrada = SaldoIdEntrada.toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      SaldoIdEntrada = SaldoIdEntrada.replace(/\./g,"");
      var Cod = Entradas[i][7];
      var Pu = Entradas[i][11].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      Pu = Pu.replace(/\./g,"");
      var Obs = Entradas[i][12];
      var Qpc = Entradas[i][8];
      var UnidadeDistribuicao = Entradas[i][14];
      var ElementoDespesa = Entradas[i][13];
      break;
    }

   }

  }   

  dadosEntradas.length = 0;
  dadosSaidas.length = 0;
  Entradas.length = 0;
  Saidas.length = 0;

  return ([SaldoTotal, IdEntrada, Cod, Pu, SaldoIdEntrada, Obs, Qpc, UnidadeDistribuicao, ElementoDespesa]);

}

function getPositiveSaldoMarcas(Linha) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaEntrada = planilha.getSheetByName("Entradas");
  var guiaSaida = planilha.getSheetByName("Saídas");

  var ultimaLinhaEntrada = guiaEntrada.getLastRow();
  var dadosEntradas = guiaEntrada.getRange(2, 1, ultimaLinhaEntrada, 16).getValues();
  var marcasEntradas = dadosEntradas.filter(function (value) {
    return value[4] == Linha;
  });

  var ultimaLinhaSaida = guiaSaida.getLastRow();
  var dadosSaidas = guiaSaida.getRange(2, 1, ultimaLinhaSaida, 17).getValues();
  var marcasSaidas = dadosSaidas.filter(function (value) {
    return value[4] == Linha;
  });

  var saldoMarcas = {};
  
  marcasEntradas.forEach(function (entrada) {
    var marca = entrada[5];
    if (!saldoMarcas[marca]) saldoMarcas[marca] = 0;
    saldoMarcas[marca] += parseFloat(entrada[10]) || 0;
  });

  marcasSaidas.forEach(function (saida) {
    var marca = saida[5];
    if (saldoMarcas[marca] !== undefined) {
      saldoMarcas[marca] -= parseFloat(saida[9]) || 0;
    }
  });

  var marcasComSaldoPositivo = [];
  for (var marca in saldoMarcas) {
    if (saldoMarcas[marca] > 0) {
      marcasComSaldoPositivo.push(marca);
    }
  }

  return marcasComSaldoPositivo;
}



function SalvarSaida(Dados) { 
  const user = LockService.getScriptLock();
  user.tryLock(10000);
  if (user.hasLock()) {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaSaida = planilha.getSheetByName("Saídas");
    var maiorId = Math.max.apply(null, guiaSaida.getRange("A2:A").getValues());
    var novoId = maiorId + 1;
    var linha = guiaSaida.getLastRow() + 1;

    var dataQuebrada = Dados.Data.split("/");
    var Ano = dataQuebrada[0];
    var Mes = dataQuebrada[1];
    var Dia = dataQuebrada[2];
    var Data = Dia + "/" + Mes + "/" + Ano;

    var DataMes = new Date(Dados.Data);
    var m = DataMes.getMonth();
    var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    var Mes = meses[m];

    guiaSaida.getRange(linha, 1).setValue(novoId);
    guiaSaida.getRange(linha, 2).setValue(Data);
    guiaSaida.getRange(linha, 3).setValue(Mes);
    guiaSaida.getRange(linha, 4).setValue(Ano);
    guiaSaida.getRange(linha, 5).setValue(Dados.Linha);
    guiaSaida.getRange(linha, 6).setValue(Dados.Marca);
    guiaSaida.getRange(linha, 7).setValue(Dados.Produto);
    guiaSaida.getRange(linha, 8).setValue(Dados.Ide);
    guiaSaida.getRange(linha, 9).setValue(Dados.Pu);
    guiaSaida.getRange(linha, 10).setValue(Dados.Qtd);
    guiaSaida.getRange(linha, 11).setValue(Dados.Valor);
    guiaSaida.getRange(linha, 12).setValue(Dados.Cod);
    guiaSaida.getRange(linha, 13).setValue(Dados.Obs);
    guiaSaida.getRange(linha, 14).setValue(Dados.Pedido);
    guiaSaida.getRange(linha, 15).setValue(Dados.Qpc);
    guiaSaida.getRange(linha, 16).setValue(Dados.UnidadeDistribuicao);
    guiaSaida.getRange(linha, 17).setValue(Dados.ElementoDespesa);

    return "SALVO COM SUCESSO!"; 
  }  
}


function PesquisarSaida(Id){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaSaida = planilha.getSheetByName("Saídas");

  var ultimaLinha = guiaSaida.getLastRow();

  var dados = guiaSaida.getRange(2, 1, ultimaLinha, 17).getValues();

  for(var i = 0; i < dados.length; i++){

    if(dados[i][0] == Id){

      var Id = dados[i][0];
      var data = new Date(dados[i][1]);

      var Dia = data.getDate();
      var Mes = data.getMonth() + 1;
      var Ano = data.getFullYear();

      var Data = Ano + "-" + Mes + "-" + Dia;

      var Linha = dados[i][4];
      var Marca = dados[i][5];
      var Produto = dados[i][6];
      var Ide = dados[i][7];

      var P = dados[i][8].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      var Pu = P.replace(/\./g,"");

      var Q = dados[i][9].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      var Qtd = Q.replace(/\./g,"");

      var V = dados[i][10].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      var Valor = V.replace(/\./g,"");

      var Cod = dados[i][11];
      var Obs = dados[i][12];

      var Pedido = dados[i][13];
      var Qpc = dados[i][14];
      var UnidadeDistribuicao = dados[i][15];
      var ElementoDespesa = dados[i][16];


      dados.length = 0;

      return ([Id, Data, Linha, Marca, Produto, Ide, Pu, Qtd, Valor, Cod, Obs, Pedido, Qpc, UnidadeDistribuicao, ElementoDespesa]);

    }

  }

  dados.length = 0;

  return "NÃO ENCONTRADO!";

}

function SaldoEditar(Dados){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaEntrada = planilha.getSheetByName("Entradas");
  var guiaSaida = planilha.getSheetByName("Saídas");

  var Linha = Dados.Linha;
  var Marca = Dados.Marca;
  var Produto = Dados.Produto;
  var IdEntrada = Dados.Ide;

  var ultimaLinha = guiaEntrada.getLastRow();
  var dadosEntradas = guiaEntrada.getRange(2,1, ultimaLinha, 16).getValues();

  var Entradas = dadosEntradas.filter(function(value, i, arr){
    return Linha == arr[i][4] && Marca == arr[i][5] && Produto == arr[i][6];
  });

  if(Entradas.length == 0){
    dadosEntradas.length = 0;
    return "NÃO TEM ENTRADA PARA ESTE PRODUTO!";
  }

  var ultimaLinha = guiaSaida.getLastRow();
  var dadosSaidas = guiaSaida.getRange(2,1, ultimaLinha, 17).getValues();

  var Saidas = dadosSaidas.filter(function(value,i,arr){
    return Linha == arr[i][4] && Marca == arr[i][5] && Produto == arr[i][6];
  });

  var QtdEntrada = 0;
  var QtdSaida = 0;

  for(var i = 0; i < Entradas.length; i++){
    if(Entradas[i][10] != ""){
      QtdEntrada = parseFloat(QtdEntrada) + parseFloat(Entradas[i][10]);
    }
  }

  for(var i = 0; i < Saidas.length; i++){
    if(Saidas[i][9] != ""){
      QtdSaida = parseFloat(QtdSaida) + parseFloat(Saidas[i][9]);
    }
  }

  var SaldoTotal = parseFloat(QtdEntrada) - parseFloat(QtdSaida);
  SaldoTotal = SaldoTotal.toLocaleString({style: 'decimal', decimal: 'pt-BR'});
  SaldoTotal = SaldoTotal.replace(/\./g,"");

  for(var i = 0; i < Entradas.length; i++){ 

    var QtdIdEntrada = Entradas[i][10];
    var SaidaIdEntrada = 0;
    var SaldoIdEntrada = 0;

    for(var l = 0; l < Saidas.length; l++){
      if(Entradas[i][0] == IdEntrada && Saidas[l][7] == IdEntrada){
        SaidaIdEntrada = parseFloat(SaidaIdEntrada) + parseFloat(Saidas[l][9]);
      }
    }

  if(parseFloat(QtdIdEntrada) > parseFloat(SaidaIdEntrada) && Entradas[i][0] == IdEntrada){
    SaldoIdEntrada = parseFloat(QtdIdEntrada) - parseFloat(SaidaIdEntrada);
    SaldoIdEntrada = SaldoIdEntrada.toLocaleString({style: 'decimal', decimal: 'pt-BR'});
    SaldoIdEntrada = SaldoIdEntrada.replace(/\./g,"");      
    break;
  }  

  }   

  dadosEntradas.length = 0;
  dadosSaidas.length = 0;
  Entradas.length = 0;
  Saidas.length = 0;

  return ([SaldoTotal, SaldoIdEntrada]);

}

function EditarSaida(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaSaida = planilha.getSheetByName("Saídas");

    var ultimaLinha = guiaSaida.getLastRow();

    var dados = guiaSaida.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dados.length; i++){

      if(dados[i][0] == Dados.Id){

        var linha = i + 2;

        var dataQuebrada = Dados.Data.split("/");
        var Ano = dataQuebrada[0];
        var Mes = dataQuebrada[1];
        var Dia = dataQuebrada[2];
        var Data = Dia + "/" + Mes + "/" + Ano;

        var DataMes = new Date(Dados.Data);
        var m = DataMes.getMonth();

        var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

        var Mes = meses[m];
        
        guiaSaida.getRange(linha, 2).setValue(Data);
        guiaSaida.getRange(linha, 3).setValue(Mes);
        guiaSaida.getRange(linha, 4).setValue(Ano);
        guiaSaida.getRange(linha, 5).setValue(Dados.Linha);
        guiaSaida.getRange(linha, 6).setValue(Dados.Marca);
        guiaSaida.getRange(linha, 7).setValue(Dados.Produto);
        guiaSaida.getRange(linha, 8).setValue(Dados.Ide);
        guiaSaida.getRange(linha, 9).setValue(Dados.Pu);
        guiaSaida.getRange(linha, 10).setValue(Dados.Qtd);
        guiaSaida.getRange(linha, 11).setValue(Dados.Valor);
        guiaSaida.getRange(linha, 12).setValue(Dados.Cod);
        guiaSaida.getRange(linha, 13).setValue(Dados.Obs);
	      guiaSaida.getRange(linha, 14).setValue(Dados.Pedido);
        guiaSaida.getRange(linha, 15).setValue(Dados.Qpc);
        guiaSaida.getRange(linha, 16).setValue(Dados.UnidadeDistribuicao);
        guiaSaida.getRange(linha, 17).setValue(Dados.ElementoDespesa);

        dados.length = 0;

        return "EDITADO COM SUCESSO!"; 

      }

    }

    dados.length = 0;

    return "ID NÃO ENCONTRADO!"; 

  }

}

function ExcluirSaida(Id){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaSaida = planilha.getSheetByName("Saídas");

    var ultimaLinha = guiaSaida.getLastRow();

    var dados = guiaSaida.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dados.length; i++){

      if(dados[i][0] == Id){

        var linha = i + 2;
        guiaSaida.deleteRow(linha);
        dados.length = 0;
        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dados.length = 0;
    return "NÃO ENCONTRADO!";

  }

}





























