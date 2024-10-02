function getPossibleWords() {
  let possibleWords = PropertiesService.getScriptProperties().getProperty('possibleWords');
  if (!possibleWords) {
    possibleWords = [
    'JLS', 'Kucmaq', 'Transforme Lojas', 'Cavazotto', 'Patrimar', 'Accord',
    'Inspirare', 'Perfil Maq', 'FPS', 'Pelo e Pele', 'Jovino', 'BOOM',
    'Laborene', 'Engservice', 'Transforme Loja', 'Platocom', 'MRSUL', 'ACR',
    'Factor', 'Ferrante', 'Pirow', 'Capanema', 'Capanema Escapamentos',
    'Transforme', 'Allkit', 'Reccato', 'Giacomini', 'Salutem', 'Unimov',
    'Proinox', 'Delta', 'Formatec', 'Metalbauer', 'Perfisa', 'Scheffer', 'Schefer',
    'Sulmóbile', 'sulmobile', 'DW', 'Dioxide', 'Motion', 'SOLUÇÃO', 'ABF', 
    'multsteel', 'pauluci', 'dynamics', 'alpasul', 'Cap', 'Fermolde', 'Carone',
    'Brasvac', 'Perfil Z', 'RDP', 'Metalmeth', 'Forte Rocha', 'Imasa', 'Scatolin',
    'Meac', 'Plastwin', 'Prolife', 'ABF ENGENHARIA', 'ACR INDUSTRIAL', 'ACO COELHO',
    'Coelho', 'ADRIMOVEIS', 'Adrimóveis', 'AG MOVELARIA', 'AG TECHNOLOGY', 'AGRIAÇO',
    'ALPA INDUSTRIAL', 'ALPA', 'ARIARTE', 'ARTE ESTOFADOS DECOR', 'ARTE Estofados', 
    'ARTE FARMA', 'ARTLINE', 'BALANCAS CAPITAL', 'BALANCAS', 'BALCONY', 'BECKHAUSER', 
    'BELEM', 'BOOM DO BRASIL', 'MET BERTIZOLA', 'Bertizola', 'BOLIS DESING', 'BOLIS', 
    'CAICARA ARTEFATOS', 'CAICARA', 'CALLI DO BRASIL', 'CALLI', 
    'CARFER', 'CARPAN', 'CASA DO BRAILLE', 'CEQUINATTO', 'C A', 
    'CMSI', 'CONSTRUTORA SAIMOR', 'Saimor', 'CORDASSO MOVEIS', 'CORDASSO', 'CORPO DOURADO', 
    'CORPO', 'CPA COLLORS', 'CPA', 'DECORMADE', 'DELLON', 'DELTA BRASIL',
    'DOMARCO', 'DUNE ORTOPEDICOS', 'DUNE', 'DW ELEVADORES', 'ECOMETAL', 
    'ECOMETH', 'EMBREPARTS', 'ENGFARMA', 'ERL SILOS', 'EUROMOVEL', 'EUROTECH', 
    'FATTORE', 'FIRENZE', 'METAL SCHILIN', 'Schilin', 'FENIX ESTOFADOS', 'Fenix',
    'FERRATI', 'FERREIRA SOLUCOES', 'FERREIRA', 'USIKRAFT', 'Horizonte', 'ARES BRASIL', 'ARES', 
    'FORTE', 'FPS BRASIL', 'FRONTALMAQ', 'FUNDICOES COLUMBIA', 'Columbia', 
    'GABINETES IMIGRANTES', 'Imigrantes', 'GAVIOLI', 'GIMAK', 'GOLD LINE', 'GOLD', 
    'GRANLUX', 'HEMF DIVISORIAS', 'HEMF', 'MONTAK', 'IAC IMPLEMENTOS', 'IAC', 'IDEAL RUPOLO', 'IDEAL', 
    'IMOL', 'IMPERIAL CUTELARIA', 'IMPERIAL', 'IROTEC', 'INOVAR',
    'JS ESTOFADOS', 'JS', 'JOMETAL', 'JOTA PR ARMARIOS', 'JOTA',
    'K2K DECORACAO', 'K2K', 'KILBRA', 'KIRIUS', 'KRONBAUER', 'HIDROPAN', 'KS INDUSTRIA', 'KS', 
    'DISTRIBUIDORA LABORENE', 'LAGUNA PORTAS', 'LAGUNA', 'LD EQUIPAMENTOS', 
    'LD', 'LIGHT DESIGN', 'LIGHT', 'LIMERFILMES', 'LINECOM', 'LINOPLAST', 'LJ FERRAMENTARIA', 'LJ', 
    'LUNASA', 'MADEIREIRA SORRISO', 'Sorriso', 'MADRESILVA', 'MARUPA', 'MARZO', 'ZM BOMBAS', 'ZM', 
    'MECAINOX', 'MECMAQ', 'METALURGICA CIRCULO', 'Circulo', 'MET. CENTRAL', 'Central',
    'MEVAL', 'MICHIBEL', 'MODILAC', 'MONT FACIL', 'MOTION ELEVADORES', 'Motion', 'MOVEIS GERMAN', 
    'German', 'MOVEIS OSTERNO', 'Osterno', 'PRIMOLAN', 'MOVEIS SCHUSTER', 'Schuster', 'MOVELTEC', 
    'MRE INDUSTRIAL', 'MRE', 'MT MOTO ACESSORIOS', 'MT', 'MULTIPET', 'MULTISTEEL', 'NEVES RATTAN', 
    'NINGER BOMBAS', 'Ninger', 'ODONTOPLAY', 'OLVEPIN', 'ORTHOFLEX', 'PARANAVAI',
    'PERFILADOS PERFISA', 'PHARMA', 'PLANCUS', 'PORTAS RECK', 'Reck', 'PRIMEINOX', 'PRODOMO', 'PRODSTEEL', 
    'PROJETO PROMOV', 'Promov', 'PROFILE', 'RBL USINAGEM', 'RBL', 'RECCATO SOLUÇÕES', 
    'RODAPAR', 'BATERMAQ', 'RONEBER', 'L A PRODUTOS CERAMICOS', 'RST INDUSTRIA', 'RST', 
    'S. A. MOVEIS', 'S. A.', 'F N S IND', 'FNS', 'SCHROEDER', 'SDS IND', 'SDS', 'SILOBRAS', 'SMA INDUSTRIAL', 
    'SMA', 'INCOMOL', 'SONETTO', 'STRONGFER', 'SUL AMERICANA', 'SULMETAL', 'SULMOBILE', 'SUZUMAK', 
    'TCHE', 'TECSILOS', 'TEMPO OPERACOES', 'TEMPO', 'TROMINK', 'TRONCOS E BALANCAS PROGRESSO', 
    'TRONCOS', 'Progrsso', 'UCHOA MOVEIS', 'UCHOA', 'UNIVERSAL OFFICE', 'Universal', 'VISCONOBRE', 'VISUAL TENDAS', 
    'VISUAL', 'VOSSER AQUECEDORES', 'VOSSER', 'WINNING PACK', 'WINNING', 'ZEEP', 'ZOBOR',
    'Ferdimat', 'Nardimaq',/*temp*/'Nardimac', 'Ls Brasil', 'Ls', 'Ampla', 'Sofisticasa', 'Horizonte',
    'Bignotto', 'Bignoto', 'São José', 'São Jose', 'Inox SJ', 'Fabbris', 'agromec', 'agromeq', 'MDM', 'mdm', 'Mega',
    'Ortometal', 'Divino', 'AW', 'Automatiza World', 'Automatiza'
  ];
    PropertiesService.getScriptProperties().setProperty('possibleWords', JSON.stringify(possibleWords));
  } else {
    possibleWords = JSON.parse(possibleWords);
  }
  return possibleWords;
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
