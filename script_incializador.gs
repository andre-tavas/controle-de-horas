meses = {
  1:'Janeiro',
  2:'Fevereiro',
  3:'Março',
  4:'Abril',
  5:'Maio',
  6:'Junho',
  7:'Julho',
  8:'Agosto',
  9:'Setembro',
  10:'Outubro',
  11:'Novembro',
  12:'Dezembro'
  }

/**
 * Inicializa e cria as abas da planilha
 */
function inicializador(){
  abaControle();
  abaAnaliseEquipes("Projetos");
  abaAnaliseEquipes("Equipes");
}

/**
 * Cria a aba de controle com os dados dos usuarios
 */
function abaControle() {
  var spreadsheet = SpreadsheetApp.openByUrl(WORKSHEET_LINK);
  var init_config = getConfigInfo();


  // Insere a aba controle
  spreadsheet.insertSheet('Controle');

  // Cria titulo da tabela de horas mensais
  spreadsheet.getRange('R2C3:R2C'+String(init_config.meses+3)).merge().activate();
  spreadsheet.getActiveRange().setValue('Horas Mensais');
  spreadsheet.getActiveRange()
    .setBackground('#543396')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(24)
    .setHorizontalAlignment('center');
  
  // Cria cabecalho da tabela de horas mensais
  spreadsheet.getRange('C3').setValue('Nome')
  .setFontWeight('bold')
  .setFontColor('BACKGROUND')
  .setHorizontalAlignment('center');

  var date = Utilities.formatDate(init_config.data_inicial, 'America/Sao_Paulo', 'dd/MM/yyyy');
  var first_month = Number(date.slice(3,5));

  for(var month = first_month; month <= first_month + init_config.meses; month++){
    if(month != first_month + init_config.meses){
    spreadsheet.getRange('R3C'+String(4 + month-first_month)).setValue(meses[month])
    .setFontWeight('bold')
    .setFontColor('BACKGROUND')
    .setHorizontalAlignment('center');
    }
    if(month != 13){
    spreadsheet.getRange('R4C'+String(4 + month-first_month)).setValue('01/' + String(month) + date.slice(5));
    }else{
      spreadsheet.getRange('R4C'+String(4 + month-first_month)).setValue('01/01/' + String(1+Number(date.slice(6))));
    }
  }

  // Coloca os nomes e cargos na tabela de horas
  var info = getUsersInfo(spreadsheet);
  for(var user = 0; user < info.nomes.length; user++){
    spreadsheet.getRange('B'+String(5 + user)).setValue(info.cargo[user]);
    spreadsheet.getRange('C'+String(5 + user)).setValue(info.nomes[user]);
  }

  // Insere a formula de horas por mes
  spreadsheet.getRange('D5').setFormula(
    "=IFERROR(INDEX(QUERY('"+BASE_DE_DADOS+"'!$A$2:$H;\"SELECT sum (D) where month(C)=\"&MONTH(D$4)-1&\" and E = '\"&$C5&\"'\");2;1);0)");
  spreadsheet.getRange('D5').activate().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('R5C4:R'+String(spreadsheet.getActiveRange().getLastRow())+'C'+String(3+init_config.meses)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); 

  // Oculta a linha e coluna auxiliares
  spreadsheet.hideColumn(spreadsheet.getRange('B1'));
  spreadsheet.hideRow(spreadsheet.getRange('C4'));

  // Cria tabela de medias e desvio padrao
  var fisrt_row = Number(spreadsheet.getRange('C5').getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation().slice(1)) + 2;
  spreadsheet.getRange('C'+String(fisrt_row)).setValue('Média geral').setFontWeight('bold');
  spreadsheet.getRange('D'+String(fisrt_row)).activate().setFormula('=AVERAGE(D5:D'+String(fisrt_row - 2)+')');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('R'+String(fisrt_row)+'C4:R'+String(fisrt_row)+'C'+String(3+init_config.meses)),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  var cargos = Array.from(new Set(info.cargo))
  for(var i = 1; i <= cargos.length; i++){
    spreadsheet.getRange('C'+String(fisrt_row + i)).setValue('Média ' + cargos[i-1]).setFontWeight('bold');
    spreadsheet.getRange('D'+String(fisrt_row + i)).activate().setFormula('=AVERAGEIF($B$5:$B$'+String(fisrt_row - 2)+';\"'+cargos[i-1]+'\";D5:D'+String(fisrt_row - 2)+')');
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('R'+String(fisrt_row + i)+'C4:R'+String(fisrt_row + i)+'C'+String(3+init_config.meses)),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }

  // Formata tabela de medias e desvio padrao
  spreadsheet.getActiveRange().getDataRegion().setNumberFormat('0.00')
  spreadsheet.getActiveRange().getDataRegion().setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.DASHED)

  // Cria titulo da tabela de verificacao
  spreadsheet.getRange('R2C'+String(init_config.meses+5)+':R2C'+String(init_config.meses+6)).merge().activate();
  spreadsheet.getActiveRange().setValue('Verificação');
  spreadsheet.getActiveRange()
    .setBackground('#543396')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(24)
    .setHorizontalAlignment('center');
  
  // Cria cabecalho da tabela de verificacao
  spreadsheet.getRange(replaceAll(spreadsheet.getActiveRange().getA1Notation(),'2','3')).activate().setValues(
    [['Nome', 'Último registro']])
    .setFontWeight('bold')
    .setFontColor('BACKGROUND')
    .setHorizontalAlignment('center');;
  
  // Coloca os nomes na tabela de verificacao
    for(var user = 0; user < info.nomes.length; user++){
    spreadsheet.getRange('R'+String(5 + user)+'C'+String(init_config.meses+5)).setValue(info.nomes[user]);
  }

  // Adiciona as formulas da tabela de verificacao
  spreadsheet.getRange('R5C'+String(init_config.meses+6)).activate().setFormula(
    "=MAX(FILTER('"+BASE_DE_DADOS+"'!B2:B;'"+BASE_DE_DADOS+"'!D2:D=C5))");
  spreadsheet.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Formata como tabela a tabela de horas
  spreadsheet.getRange('C3').activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getActiveRange().applyRowBanding();
  var banding = spreadsheet.getActiveRange().getBandings()[0];
  banding.setHeaderRowColor('#543396')
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#d0c3f1')
  .setFooterRowColor(null);

  // Coloca bordas na tabela de horas
  spreadsheet.getActiveRange().setBorder(false,false,false,false,false,true,"black",SpreadsheetApp.BorderStyle.DASHED);

  // Formata como tabela a tabela de verificacao
  spreadsheet.getRange('R3C'+String(init_config.meses+5)).activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getActiveRange().applyRowBanding();
  banding = spreadsheet.getActiveRange().getBandings()[0];
  banding.setHeaderRowColor('#543396')
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#d0c3f1')
  .setFooterRowColor(null);

  // Coloca bordas na tabela de verificacao
  spreadsheet.getActiveRange().setBorder(false,false,false,false,false,true,"black",SpreadsheetApp.BorderStyle.DASHED);

  // Altera o formato e centraliza coluna das datas
  spreadsheet.getRange('R5C'+String(init_config.meses+6)).activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getActiveRange().setNumberFormat('dd/MM/yyyy')
    .setHorizontalAlignment('center');

  // Formatacao condicional na coluna das datas
  var conditionalFormatRules = spreadsheet.getSheetByName('Controle').getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getRange('R5C'+String(init_config.meses+6)+':R'+String(4+info.nomes.length)+'C'+String(init_config.meses+6))])
    .whenDateBefore(SpreadsheetApp.RelativeDate.PAST_WEEK)
    .setBackground('#F4CCCC')
    .build());
  spreadsheet.getSheetByName('Controle').setConditionalFormatRules(conditionalFormatRules);

  // Centraliza o texto das horas
  spreadsheet.getRange('D5').activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getActiveRange().setHorizontalAlignment('center');

  // Altera o tamanho das margens
  spreadsheet.setColumnWidth(1, 25);
  spreadsheet.setColumnWidth(3,170);
  spreadsheet.setColumnWidth(init_config.meses+4,30);
  spreadsheet.setColumnWidth(init_config.meses+5,170);
  spreadsheet.setColumnWidth(init_config.meses+6,150);
  spreadsheet.setRowHeight(1, 15);
  
  // Desativa as linhas de grade
  spreadsheet.getSheetByName('Controle').setHiddenGridlines(true)
}

/**
 * Configura a aba tipoEquipe que analisa
 * tipoEquipe pode receber Projetos ou Equipes 
 */
function abaAnaliseEquipes(tipoEquipe){
  var spreadsheet = SpreadsheetApp.openByUrl(WORKSHEET_LINK);
  var info = getUsersInfo(spreadsheet);
  var worksheet = spreadsheet.getSheetByName(tipoEquipe);

  // Armazena os dados necessarios para configurar a aba
  var equipes = getUnique(info[tipoEquipe]);
  var cargos = getUnique(info.cargo);


  // Adiciona colunas nas tabelas caso seja necessário
  var duracao = getConfigInfo().duracao_projetos;
  worksheet.getRange('C21').activate();
  var numColunas = worksheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).getNumColumns() - 1;
  var colsToAdd = duracao - numColunas;
  var secondPartFisrtCol = 5 + numColunas;

  if(duracao != numColunas){
    // Insere as colunas na tabela da direita
    worksheet.insertColumnsAfter(secondPartFisrtCol + numColunas, colsToAdd);

    // Completa a tabela de media mensal
    worksheet.getRange(20,secondPartFisrtCol + numColunas,2).autoFill(worksheet.getRange(20,secondPartFisrtCol + numColunas,2,1 + colsToAdd), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    // Completa a tabela por projeto
    worksheet.getRange(24,secondPartFisrtCol + numColunas,3).autoFill(worksheet.getRange(24,secondPartFisrtCol + numColunas,3,1 + colsToAdd), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    // Insere as colunas na tabela da esquerda
    worksheet.insertColumnsAfter(2 + numColunas, colsToAdd);

    // Completa a tabela de media mensal
    worksheet.getRange(20, 2 + numColunas,2).autoFill(worksheet.getRange(20, 2 + numColunas,2,1 + colsToAdd), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    // Completa a tabela por projeto
    worksheet.getRange(24, 2 + numColunas,3).autoFill(worksheet.getRange(24, 2 + numColunas,3,1 + colsToAdd), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    // Completa a linha auxiliar dos cargos
    worksheet.getRange(3, 2 + numColunas).autoFill(worksheet.getRange(3, 2 + numColunas,1,1 + colsToAdd), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    // Altera a funcao de media
    worksheet.getRange(21, duracao + 3).setFormulaR1C1("=AVERAGE(R[0]C[-"+duracao+"]:R[0]C[-1])");
    worksheet.getRange(26, duracao + 3).setFormulaR1C1("=AVERAGE(R[0]C[-"+duracao+"]:R[0]C[-1])");
    worksheet.getRange(26, secondPartFisrtCol + colsToAdd + duracao + 1).setFormulaR1C1("=AVERAGE(R[0]C[-"+duracao+"]:R[0]C[-1])");
    worksheet.getRange(21, secondPartFisrtCol + colsToAdd + duracao + 1).setFormulaR1C1("=AVERAGE(R[0]C[-"+duracao+"]:R[0]C[-1])");

  }
  if(duracao <=4){
    worksheet.setColumnWidth(4+numColunas,22 + (5 - duracao)*100);
  }
  
  secondPartFisrtCol = secondPartFisrtCol + colsToAdd;

  // Coloca o nome dos meses na primeira linha das tabelas
  for(var i = 0; i < numColunas + colsToAdd; i++){
    var numMes = worksheet.getRange('R25C' + (3 + i)).getValue();
    worksheet.getRange('R24C' + (3 + i)).setValue(meses[numMes]);
    worksheet.getRange('R24C' + (secondPartFisrtCol + 1 + i)).setValue(meses[numMes]);

    worksheet.getRange('R20C' + (3 + i)).setValue(meses[numMes]);
    worksheet.getRange('R20C' + (secondPartFisrtCol + 1 + i)).setValue(meses[numMes]);
  }


  // Coloca os cargos
  for(var i = 0; i < cargos.length; i++){
    // Adiciona o cargo
    worksheet.getRange(1, 3+i)
    .setValue(cargos[i])

    // Coloca o checkbox
    worksheet.getRange(2, 3+i).insertCheckboxes();

    // Altera a formatacao
    worksheet.getRange(1, 3, 2).copyFormatToRange(worksheet, 3 + i, 3 + i, 1, 2);
  }

  
  // Preenche as equipes
  worksheet.getRange(26,2,equipes.length)
  .setValues(equipes.map(
    function(value){
      return [value]
      }));

  worksheet.getRange(26,secondPartFisrtCol,equipes.length)
  .setValues(equipes.map(
    function(value){
      return [value]
      }));


  // Formata a primeira coluna das tabelas
  worksheet.getRange('B26').activate();
  worksheet.getRange('B26').copyFormatToRange(worksheet, 2, 2, 26, 25+equipes.length);

  worksheet.getRange(26,secondPartFisrtCol).activate();
  worksheet.getRange(26,secondPartFisrtCol).copyFormatToRange(worksheet, secondPartFisrtCol, secondPartFisrtCol, 26, 25+equipes.length);


  // Replica as formulas da primeira linha
  worksheet.getRange(26,3).activate();
  worksheet.getSelection()
  .getNextDataRange(SpreadsheetApp.Direction.NEXT)
  .autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  worksheet.getRange(26,secondPartFisrtCol + 1).activate();
  worksheet.getSelection()
  .getNextDataRange(SpreadsheetApp.Direction.NEXT)
  .autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);


  // Oculta as linhas auxiliares
  worksheet.hideRow(worksheet.getRange("A3"));
  worksheet.hideRow(worksheet.getRange("A25"));
}

/**
 * Recebe uma lista em que cada elemento é uma string separando
 * elementos por ", " e retorna uma lista com elementos unicos
 */
function getUnique(lista){
  var unique_values = [];

  for(var i=0; i < lista.length; i++){
    var sublist = lista[i].split(", ");
    for(var j = 0; j < sublist.length; j++){
      if(sublist[j] != '' && ! unique_values.includes(sublist[j])){
        unique_values = unique_values.concat(sublist[j]);
      }
    }
  }
  return unique_values;
}

/**
 * Retorna as informacoes da aba de configuracoes
 */
function getConfigInfo(){
  // Ativa o arquivo utilizado
  var sheet = SpreadsheetApp.openByUrl(WORKSHEET_LINK).getSheetByName(CONFIG);

  var obj = {}

  obj['data_inicial'] = sheet.getRange('C3').getValue();

  obj['meses'] = sheet.getRange('C5').getValue();

  obj['duracao_projetos'] = sheet.getRange('C7').getValue();

  obj['inicio_projetos'] = sheet.getRange('C9').getValue();

  obj['duracao_atividade'] = sheet.getRange('F3').getValue();

  return obj;
}

/**
 * Substitui todas os caracteres em uma string
 */
function replaceAll(str, find, replace) {
  return str.replace(new RegExp(find, 'g'), replace);
}

/**
 * Atualiza as funções da aba quando a informação está em uma coluna diferente
 */
function updateSheet(teamColumn, aba){
  // Altera a fórmula na tabela da esquerda
  var sheet = SpreadsheetApp.openByUrl(WORKSHEET_LINK).getSheetByName(aba);
  sheet.getRange('C26').activate().setFormula('=DSUM(QUERY(\''+BASE_DE_DADOS+'\'!$A$1:$H;\"SELECT D,E where month(C)=\"&C$25-1&\" and '+teamColumn+'=\'\"&$B26&\"\'\");\"quantity\";{\"user\";QUERY('+USERS_INFO+'!$A$2:$E;\"select A where (E=\'\"&$C$3&\"\') or (E=\'\"&$D$3&\"\') or (E=\'\"&$E$3&\"\') or (E=\'\"&$F$3&\"\')\")})');

  var numColunas = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).getNumColumns() - 1;
  sheet.getRange(26,3,1,numColunas).activate();
  var destination = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN);
  sheet.getRange('C26').copyTo(destination);
  

  // Altera a formula na tabela da direita
  var secondPartFisrtCol = 6 + numColunas;
  sheet.getRange(26,secondPartFisrtCol).activate().setFormula('=DAVERAGE(QUERY(\''+BASE_DE_DADOS+'\'!$A$1:$H;\"select E,sum(D) where month(C)=\"&J$25-1&\" and '+teamColumn+'=\'\"&$I26&\"\' group by E\");\"sum quantity\";{\"user\";QUERY('+USERS_INFO+'!$A$2:$E;\"select A where (E=\'\"&$C$3&\"\') or (E=\'\"&$D$3&\"\') or (E=\'\"&$E$3&\"\') or (E=\'\"&$F$3&\"\')\")})');
  sheet.getRange(26,secondPartFisrtCol,1,numColunas).activate();
  destination = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN);
  sheet.getRange(26,secondPartFisrtCol).copyTo(destination);
  
}

/**
 * Atualiza as funções de todas as abas
 */
function updateFunctions(){
  var spreadsheet = SpreadsheetApp.openByUrl(WORKSHEET_LINK);
  var settings_sheet = spreadsheet.getSheetByName(CONFIG);
  var teamColumn = settings_sheet.getRange('F5').getValue();

  updateSheet(teamColumn, 'Equipes');
  updateSheet(teamColumn, 'Projetos');

  // Altera a aba de controle
  var abaControle = spreadsheet.getSheetByName('Controle');
  abaControle.getRange('D5').activate().setFormula('=IFERROR(INDEX(QUERY(\''+BASE_DE_DADOS+'\'!$A$2:$H;\"SELECT sum (D) where month(C)=\"&MONTH(D$4)-1&\" and E = \'\"&$C5&\"\'\");2;1);0)')
  
}
