function onEdit(e) {
  Logger.log('Edit event triggered');
  var range = e.range; // Captura o intervalo que foi editado
  setBackgroundOnWrite(range);
  setBackgroundOnClear(range);
}

function setBackgroundOnClear(range) {
  var sheet = range.getSheet(); // Obtém a planilha onde a edição ocorreu
  var backgrounds = range.getBackgrounds(); // Obtém as cores de fundo das células no intervalo editado
  var values = range.getValues(); // Obtém os valores das células no intervalo editado

  // Itera sobre cada célula no intervalo editado
  for (var i = 0; i < range.getNumRows(); i++) {
    for (var j = 0; j < range.getNumColumns(); j++) {
      var cell = sheet.getRange(range.getRow() + i, range.getColumn() + j);
      var newValue = values[i][j]; // Novo valor da célula
      var currentBackground = backgrounds[i][j]; // Cor de fundo atual da célula

      // Verifica se a célula está vazia agora e a cor de fundo era amarela (#ffc000) ou vermelha
      if (!newValue && (currentBackground === '#ffc000' || currentBackground === '#ff0000')) {
        cell.setBackground('#00b050'); // Muda a cor de fundo para verde
      }
    }
  }
}

function setBackgroundOnWrite(range) {
  var sheet = range.getSheet(); // Obtém a planilha onde a edição ocorreu
  var backgrounds = range.getBackgrounds(); // Obtém as cores de fundo das células no intervalo editado
  var values = range.getValues(); // Obtém os valores das células no intervalo editado

  // Itera sobre cada célula no intervalo editado
  for (var i = 0; i < range.getNumRows(); i++) {
    for (var j = 0; j < range.getNumColumns(); j++) {
      var cell = sheet.getRange(range.getRow() + i, range.getColumn() + j);
      var newValue = values[i][j]; // Novo valor da célula
      var currentBackground = backgrounds[i][j]; // Cor de fundo atual da célula

      // Verifica se a célula tem um novo valor e a cor de fundo era verde
      if (newValue && currentBackground === '#00b050') {
        cell.setBackground('#ffc000'); // Muda a cor de fundo para amarelo
      }
    }
  }
}
