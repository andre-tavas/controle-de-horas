function getAllUsersData(){
  var init_time = new Date();

  var spreadsheet = SpreadsheetApp.openByUrl(WORKSHEET_LINK);
  var destination_sheet = spreadsheet.getSheetByName(BASE_DE_DADOS)

  var info = getUsersInfo(spreadsheet);
  var numUsers = info.nomes.length;

  for(var user = 0; user < numUsers; user++){
    fetchFromTGGL(destination_sheet, info, user);
  }  

  destination_sheet.getRange('A2:I').removeDuplicates([3,5]);
  destination_sheet.getRange('A2:I').sort(3);

  checkOutliers(destination_sheet,spreadsheet);

  var end_time = new Date();
  console.log(init_time)
  console.log(end_time)
}

function fetchFromTGGL(sheet, user_info, user_num) {
  var email = user_info.email[user_num];
  var workspace = user_info.workspace[user_num];
  var api = user_info.api[user_num];

  console.log(user_info.nomes[user_num])

  // Fetch dos dados do toggl para pegar o numero de paginas
  var response = UrlFetchApp.fetch(
    getUrl(email,workspace,1), {
      headers: { Authorization: toEcodedHeader(api) },
  });

  console.log(JSON.parse(response.getContentText())['total_count'])

  var pages = numPages(response);

  for(var j = 1; j <= pages; j++){

    // Fetch para pegar os dados da pagina j
    var response = UrlFetchApp.fetch(
      getUrl(email,workspace,j), {
        headers: { Authorization: toEcodedHeader(api) },
    });

    // Resposta do toggl armazenada em .data
    var dataArr = JSON.parse(response.getContentText()).data;

    // Transforma payload da api em dados que serao usados
    var parsedResponse = dataArr.map(parseTogglDataRowForInvoice);

    // Define o cabecalho
    var row = parsedResponse[0];
    try{
    Object.keys(row).forEach((key, i) => {
    var range = sheet.getRange(1, i + 1, 1, 1);
    range.setValue(key);
    });
    }catch(e){
      console.log(e)
    }

    // Registra os dados nas linhas
    writeRowArrToSheet(parsedResponse,sheet);}
}

function toEcodedHeader(api){
  var blob_string = `${api}:api_token`;
  return Utilities.base64Encode(blob_string);
}

// Get URL to send our request to. 
const getUrl = (email,workspace,page) => {
  // Pega a janela de tempo da requisicao
  const date = new Date();
  date.setDate(date.getDate() - DAYS_2_LOOKUP);
  const dateString = `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;

  // Retorna o link para a requisicao
  return 'https://api.track.toggl.com/reports/api/v2/details?user_agent='+email
          +'&workspace_id='+workspace
          +'&display_hours=decimal'
          +'&since='+dateString
          +'&page='+page;
};

// Converte a durcao de millisegundos para horas e coloca no formato 2000 (ao inves de 2,000)
const convertDuration = (duration) => {
  const HOUR = 1000*3600;
  return (duration / HOUR).toFixed(2).replace('.',',');
};

// Converts dateTime to 2020-05-05 modify this to get a different output
// Converte formato da data para 02/05/2021 19:30
const convertDate = (dateTime) => {
  const date = new Date(dateTime);
  return date.toLocaleString('pt-BR');
};

// Remove os espacos vazios no final da string da descricao
const removeBlank = (description) => {
return description.trim();
};

// Faz ajustes no payload para ter somente os dados desejados e no formato desejado
const parseTogglDataRowForInvoice = (togglObj) => {
  
  return Object.keys(togglObj).reduce((acc, key) => {

    // Exclui campos desnecessarios
    if (FIELDS_2_IGNORE.includes(key)) return acc;

    // Converte a durcao de millisegundos para horas e renomeia a chave
    if (key === "dur") {
    return { ...acc, quantity: convertDuration(togglObj[key]) };
    }

    // Remove os espaços vazios ao final das descricoes
    if (key === "description") {
    return { ...acc, description: removeBlank(togglObj[key]) };
    }
    
    // Converte a data para o formato 02/05/2021 19:30
    if (key === "start") {
    return { ...acc, date: convertDate(togglObj[key]) };
    }

    // Converte a data para o formato 02/05/2021 19:30
    if (key === "end") {
    return { ...acc, date: convertDate(togglObj[key]) };
    }

    return { ...acc, [key]: togglObj[key] };
    }, {});
};

// https://mashe.hawksey.info/2020/04/google-apps-script-patterns-writing-rows-of-data-to-google-sheets-the-v8-way/#gref
const writeRowArrToSheet = (arr,sheet) => {
  //console.log('arr:\n'+arr)

  //const sheet = SpreadsheetApp.getActiveSheet();
  // getting our headers
  const heads = sheet.getDataRange().offset(0, 0, 1).getValues()[0];
  //console.log('Heads:\n' + heads)

  // convert object data into a 2d array
  const tr = arr.map((row) => heads.map((key) => row[String(key)] || ""));
  //console.log('tr:\n'+1);

  // write result
  try{
  sheet
  .getRange(sheet.getLastRow() + 1, 1, tr.length, tr[0].length)
  .setValues(tr);
  }catch(e){
    console.log(tr)
  }
};

/**
 * Retorna as informacoes das pessoas da organização
 */
function getUsersInfo(spreadsheet){
  // Ativa a aba da planilha com as informacoes
  var data = spreadsheet.getSheetByName(USERS_INFO).getDataRange().getValues();
  
  // Objeto de que armazena os dados
  var info = {};

  // Armazena o nome dos membros
  info['nomes'] = data.slice(1).map(function(value){ return value[0]; });

  // Armazena o workspace dos membros
  info['workspace'] = data.slice(1).map(function(value){ return value[1]; });

  // Armazena o workspace dos membros
  info['api'] = data.slice(1).map(function(value){ return value[2]; });

  // Armazena o workspace dos membros
  info['email'] = data.slice(1).map(function(value){ return value[3]; });

  // Armazena o cargo dos membros
  info['cargo'] = data.slice(1).map(function(value){ return value[4]; });

  // Armazena os projetos
  info['Projetos'] = data.slice(1).map(function(value){ return value[5]; });

  // Armazena as equipes internas
  info['Equipes'] = data.slice(1).map(function(value){ return value[6]; });

  return info;
};

/**
 * Retorna o numero de paginas que possui a resposta do toggl
 */
function numPages(response){
  const dataArr = JSON.parse(response.getContentText());

  //console.log(Number((dataArr['total_count']/50).toFixed(0)))
  var pages = Math.ceil(Number((dataArr['total_count']/50)))

  return pages;
}

/**
 * Envia email para o usuário quando o tempo registrado no toggl estiver errado
 */
function checkOutliers(destination_sheet,spreadsheet){
  var data = destination_sheet.getDataRange().getValues();
  var tempo = getConfigInfo().duracao_atividade;
  var row = 0;

  data.map(
    function(value){
      row ++;
      if(value[3] > tempo){
        console.log(value);
        var verificacao = destination_sheet.getRange(row,9);
        
        if(verificacao.getValue() == "Não verificado" && isEmailOn(spreadsheet)){
          sendEmail(getEmail(spreadsheet, value[4]),value);
        };
        if(verificacao.isBlank()){
          if(isEmailOn(spreadsheet)){ 
            sendEmail(getEmail(spreadsheet, value[4]),value);
            }
          verificacao.setValue("Não verificado");
        };
      };
    }
  )

}

/**
 * Envia email
 */
function sendEmail(destinatario,registro){
  var data = new Date(registro[2]);

  var texto = '<p>Olá!</p>'
                +'<p>No dia '
                +"<strong>" + data.getDate() +'/'+data.getMonth()+'/'+data.getFullYear() + "</strong>"
                +' às '
                + "<strong>" + (data.getHours() + 1) +':'+ data.getMinutes() + "</strong>"
                + ' horas você registrou a seguinte atividade: '
                +"<p><strong>" + registro[1] + "</strong></p>"
                +' com a duracao de '
                +"<strong>"+ registro[3] +" horas.</strong>"
                +'</p><p>Caso não esteja correto, busque a linha deste registro na aba '
                +'<strong>'+BASE_DE_DADOS+'</strong> na '
                +'<a href="' + WORKSHEET_LINK + '">planilha de controle de horas</a>.'
                + ' Para isso, basta apertar CTRL + F para e colar o código '+ registro[0]
                +' . Ao encontrar o registro, delete a sua linha e em seguida faça o registro manual no Toggl'
                +' com a duração correta, a descrição da atividade e a data de quando ela ocorreu e as outras informações.<\p>'
                +'<p>Caso o registro esteja correto, encontre a linha do registro na planilha usando o'
                + 'código e substitua o valor '
                +'"<b>Não verificado</b>" por "<b>Verificado</b>".</p>';

  GmailApp.sendEmail(destinatario, SUBJECT,"",{"htmlBody":texto});
}

/**
 * Retorna o email de um usuario com base em seu nome
 */
function getEmail(spreadsheet, user){
  var info = getUsersInfo(spreadsheet);
  var index = info['nomes'].indexOf(user);

  return info.email[index];
}

/**
 * Verifica se a funcionalidade de enviar email esta ativada
 */
function isEmailOn(spreadsheet){
  return spreadsheet.getSheetByName(CONFIG).getRange('I3').getValue() == 'Ativado';
}
