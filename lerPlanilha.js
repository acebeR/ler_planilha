
var Excel = require('exceljs');
var tarefaPlanilha = {};
var fs = require('fs');


async function processLineByLine() {
  let workbook = new Excel.Workbook();
  return await workbook.xlsx.readFile('Organizador de Eventos.xlsx');
}

async function init() {
  var planilha = await processLineByLine();
  organizaPlanilha(planilha);
}
init();

  function organizaPlanilha(planilha){
    var eventos = [];
    var texto = "Eventos";
    // console.log(texto);
    var abaDaVez = planilha.getWorksheet(texto);
    abaDaVez.eachRow(function(row, rowNumero) {
      row.eachCell(function(cell, cellNumero) {
        var json = {};
        if(row.number >= 12 && cellNumero >= 3 && cellNumero <= 7){
          // console.log(rowNumero, cellNumero);
          // console.log("-------------");
          // console.log(cell.value);
          switch (cellNumero) {
            case 3:
              json.data = cell.value;
              break;
            case 4:
              json.acao  = cell.value;
              break;
            case 5:
              json.data_conclusao =  cell.value;
              break;
            case 6:
              json.obs =  cell.value;
              break;
            case 7:
              json.status =  cell.value;
              break;
            default:
              console.log("Coluna n'ao mapeada!");
          }
        }
        eventos.push(json);
      }); 
    });
    console.log(eventos);
    criarArquivoJsonPlanilha(eventos);
  }

  function criarArquivoJsonPlanilha(json){
    fs.writeFile("saidaJson.txt",JSON.stringify(json), function(erro) {
      if(erro) {
          throw erro;
      }
      console.log("Arquivo json planilha salvo!");
    }); 
  }


