function onEdit(e) {
  Logger.log('Edit event triggered');
  var range = e.range; // Captura o intervalo que foi editado
  setBackgroundOnWrite(range);
  setBackgroundOnClear(range);
}

function setBackgroundOnClear(range) {
  const sheet = range.getSheet();
  const backgrounds = range.getBackgrounds();
  const values = range.getValues();
  const newBackgrounds = [];

  for (let i = 0; i < range.getNumRows(); i++) {
    const rowBackgrounds = [];
    for (let j = 0; j < range.getNumColumns(); j++) {
      if (!values[i][j] && (backgrounds[i][j] === '#ffc000' || backgrounds[i][j] === '#ff0000')) {
        rowBackgrounds.push('#00b050'); // Green for cleared cells
      } else {
        rowBackgrounds.push(backgrounds[i][j]); // Keep the original background
      }
    }
    newBackgrounds.push(rowBackgrounds);
  }

  range.setBackgrounds(newBackgrounds); // Batch update
}

function setBackgroundOnWrite(range) {
  const sheet = range.getSheet();
  const backgrounds = range.getBackgrounds();
  const values = range.getValues();
  const newBackgrounds = [];

  for (let i = 0; i < range.getNumRows(); i++) {
    const rowBackgrounds = [];
    for (let j = 0; j < range.getNumColumns(); j++) {
      if (values[i][j] && backgrounds[i][j] === '#00b050') {
        rowBackgrounds.push('#ffc000'); // Yellow for new values
      } else {
        rowBackgrounds.push(backgrounds[i][j]); // Keep the original background
      }
    }
    newBackgrounds.push(rowBackgrounds);
  }

  range.setBackgrounds(newBackgrounds); // Batch update
}

