
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
  var ilha = [];
  for(var k = 1; k <= 4; k++){
    var texto = "Eventos" + k;
    console.log(texto);
    // var ilhaDaVez = planilha.getWorksheet(texto);
    // ilhaDaVez.eachRow(function(row, rowNumero) {
    //   var i = 0;
    //   var linha = {};
    //   row.eachCell(function(cell, cellNumero) {
    //     if(row.number >= 7){
    //       i++;
    //       switch (i) {
    //         case 1:
    //           linha.id = cell.value;
    //           break;
    //         case 2:
    //           linha.itemCatalogo = cell.value;
    //           break;
    //         case 3:
    //           linha.atividade = cell.value;
    //           break;
    //         case 4:
    //           linha.data = cell.value;
    //           break;
    //         case 5:
    //           linha.hst = cell.value;
    //           break;
    //         case 6:
    //           linha.responsavel = cell.value;
    //           break;
    //         default:
    //           console.log("Coluna n'ao mapeada!");
    //       }
    //     }
    //   });
    //   if(i > 3){
    //     ilha.push(linha);
    //   }
    // });
    // var textoIlha = 'ilha' + k;
    // tarefaPlanilha[textoIlha] = ilha;
    // criarArquivoJsonPlanilha(tarefaPlanilha);
  }

  function criarArquivoJsonPlanilha(json){
    fs.writeFile("jsonPlanilha.txt",JSON.stringify(json), function(erro) {
      if(erro) {
          throw erro;
      }
      console.log("Arquivo json planilha salvo!");
    }); 
  }
}

