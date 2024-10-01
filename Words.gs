function getPossibleWords() {
  return ['words'];
}

function getAllPossibleWords(possibleWords) {
  return [
    ...possibleWords, 
    ...possibleWords.map(word => word.toLowerCase()), 
    ...possibleWords.map(word => word.toUpperCase()),
    ...possibleWords.map(word => capitalizeFirstLetter(word))
  ];
}

function capitalizeFirstLetter(word) {
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

function getPossiblePrefixes() {
  return ['Imp', 'Impl', 'Implantação', 'Implantacao', 'imp', 'impl', 'implantação', 'implantacao', 'Imp.', 'impl.', 'Impl.', 'Imp.', 'Imp/', 'impl/', 'Impl/'];
}
