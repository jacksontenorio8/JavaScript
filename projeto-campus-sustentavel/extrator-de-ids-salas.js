function organizarIdSalas() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaLinha = aba.getLastRow();
  
  // Se a planilha estiver vazia abaixo da linha 4, para a execução
  if (ultimaLinha < 4) return;

  // 1. Pega os dados da coluna G (coluna 7), começando na linha 4
  var intervaloG = aba.getRange(4, 7, ultimaLinha - 3, 1);
  var valoresG = intervaloG.getValues();
  
  var listaNovosIDs = [];
  var listaTextosLimpos = [];

  // 2. Processa cada linha
  for (var i = 0; i < valoresG.length; i++) {
    var textoOriginal = valoresG[i][0].toString();
    
    if (textoOriginal.includes(" - ")) {
      // Divide o texto pelo separador " - "
      var partes = textoOriginal.split(" - ");
      
      // O ID é o último elemento da lista
      var id = partes.pop(); // .pop() remove o último e guarda na variável
      
      // O texto limpo é o que sobrou (juntamos novamente)
      var textoLimpo = partes.join(" - ");
      
      listaNovosIDs.push([id]);
      listaTextosLimpos.push([textoLimpo]);
    } else {
      // Caso a célula não siga o padrão, mantém como está
      listaNovosIDs.push([""]); 
      listaTextosLimpos.push([textoOriginal]);
    }
  }

  // 3. Escreve os IDs na coluna E (coluna 5)
  aba.getRange(4, 5, listaNovosIDs.length, 1).setValues(listaNovosIDs);
  
  // 4. Atualiza a coluna G (coluna 7) removendo o ID de lá
  aba.getRange(4, 7, listaTextosLimpos.length, 1).setValues(listaTextosLimpos);
  
  SpreadsheetApp.getUi().alert("Processo concluído! IDs movidos para a coluna E.");
}

/**
 * Cria um menu no topo para você rodar o script com um clique
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(' Minhas Automações')
      .addItem('Mover IDs agora', 'organizarIdSalas')
      .addToUi();
}


