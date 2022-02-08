/**
 * CONFIGURAR A AUTOMAÇÃO
 * - Coloque o link do arquivo em WORKSHEET_LINK
 * - Caso necessário, altere o nome de alguma das abas
 * - Coloque a quantidade de dias a serem importados do toggl
 * - Caso queira, altere o assunto do email de alerta para atividades suspeitas
 * - Clique em "Salvar projeto"
 * - Em "Acionadores", adicione um acionador que execute a função "getAllUsersData" diariamente
 */


// Link da planilha de controle de horas
const WORKSHEET_LINK = 'https://docs.google.com/spreadsheets/d/1kE5G6uHdg8cpJoO2tLHtZNpv5K77PECEwbAWK-EREKk/edit#gid=2001691719';

// Nome da aba com as informações dos usuarios
const USERS_INFO = 'Usuários';

// Nome da aba onde os dados do toggl serao registrados
const BASE_DE_DADOS = 'Base de dados';

// Nome da aba de configurações
const CONFIG = 'Confirgurações';

// Qtde. de dias a serem importados antes da data atual
const DAYS_2_LOOKUP = 40;

// Assunto do email para registros muito altos
const SUBJECT = "Alerta registro no toggl";

// Campos que não são importados do toggl
const FIELDS_2_IGNORE = ["pid","tid","uid","updated","use_stop","project_color","project_hex_color","task","billable","is_billable","cur"];

/**
 * CAMPOS QUE SERÃO USADOS:
 * -id
 * -description
 * -date
 * -dur
 * -user
 * -client
 * -projet
 * -tags
 */
