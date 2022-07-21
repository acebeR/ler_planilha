# ler_planilha
Objetivo: ler planilhas do exel com mais de uma aba e Retornar um Json.
</br></br></br>
<h1>Para Rodar o projeto</h1>
<p>Click duas vezes no arquivo start.exe</p>
</br></br>
<h1>Tecnologias</h1>
<p>Node.js</p>
<p>JavaScripty</p>
</br></br>
<h1>Código Principal</h1>
    var planilha = await processLineByLine();
    var abaDaVez = planilha.getWorksheet(texto);
    abaDaVez.eachRow(function(row, rowNumero) {
      console.log(row);
    });
</br>
<h1>Dependências</h1>
  "dependencies": {
    "exceljs": "^4.2.1",
    "js-xlsx": "^0.8.22",
    "read-excel-file": "^5.4.0",
    "workbook": "^1.1.3",
    "dotenv": "^8.2.0"
  }
</br>
<h1>Para Criar Arquivo</h1>
  function criarArquivoJsonPlanilha(json){
    fs.writeFile("saidaJson.txt",JSON.stringify(json), function(erro) {
      if(erro) {
          throw erro;
      }
      console.log("Arquivo json planilha salvo!");
    }); 
  }