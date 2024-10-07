// Configurações iniciais
const CLIENT_ID = '###';
const ORG_ID = '###';
const CLIENT_SECRET = '##'; //substituir pelos dados da api e org
const UMAPI_URL = 'https://usermanagement.adobe.io/v2/usermanagement/users/';
const MAX_RETRIES = 4;
const TIMEOUT = 120.0;
const RANDOM_MAX = 5;
const FIRST_DELAY = 3;

function getToken() {
  var url = 'https://ims-na1.adobelogin.com/ims/token/v3';
  var payload = {
    'grant_type': 'client_credentials',
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET,
    'scope': 'openid,AdobeID,user_management_sdk'
  };

  var options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  var accessToken = json.access_token;

  return accessToken;
}

const ACCESS_TOKEN = getToken();

function getUsersInOrg() {
  let page_index = 0;
  let url = UMAPI_URL + ORG_ID + '/' + page_index;
  let method = 'GET';
  let done = false;
  let users_list = [];

  while (!done) {
    let response = makeCall(method, url);
    if (response && response.lastPage) {
      users_list = users_list.concat(response.users);
      done = true;
    } else if (response) {
      users_list = users_list.concat(response.users);
      page_index += 1;
      url = UMAPI_URL + ORG_ID + '/' + page_index;
    } else {
      done = true;
    }
  }
  return users_list;
}

function makeCall(method, url, body = {}) {
  let retry_wait = 0;
  let headers = {
    'Accept': 'application/json',
    'x-api-key': CLIENT_ID,
    'Authorization': 'Bearer ' + ACCESS_TOKEN
  };
  if (Object.keys(body).length !== 0) {
    headers['Content-type'] = 'application/json';
    body = JSON.stringify(body);
    method = 'POST';
  }

  for (let num_attempt = 1; num_attempt <= MAX_RETRIES; num_attempt++) {
    try {
      console.log(`Calling ${method} ${url}\n${body}`);
      let options = {
        method: method,
        headers: headers,
        payload: body,
        muteHttpExceptions: true
      };
      let response = UrlFetchApp.fetch(url, options);
      let statusCode = response.getResponseCode();
      if (statusCode === 200) {
        return JSON.parse(response.getContentText());
      } else if ([429, 502, 503, 504].includes(statusCode)) {
        console.log(`UMAPI timeout... (code ${statusCode} on try ${num_attempt})`);
        if (retry_wait <= 0) {
          let delay = Math.floor(Math.random() * RANDOM_MAX);
          retry_wait = (Math.pow(2, num_attempt - 1) * FIRST_DELAY) + delay;
        }
        let retryAfter = response.getHeaders()['Retry-After'];
        if (retryAfter) {
          retry_wait = parseInt(retryAfter) + 1;
        }
      } else {
        console.log(`Unexpected HTTP Status: ${statusCode}: ${response.getContentText()}`);
        return;
      }
    } catch (e) {
      console.log(`Exception encountered:\n ${e}`);
      return;
    }
    if (num_attempt < MAX_RETRIES && retry_wait > 0) {
      console.log(`Next retry in ${retry_wait} seconds...`);
      Utilities.sleep(retry_wait * 1000);
    }
  }
  console.log(`UMAPI timeout... giving up after ${MAX_RETRIES} attempts.`);
}

function main_adobe() {
  let users = getUsersInOrg();
  if (users) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Contas_Adobe');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Contas_Adobe');
    }
    sheet.clear(); // Limpa a planilha antes de inserir novos dados
    sheet.appendRow(['ID', 'Email', 'Groups']); // Cabeçalhos das colunas

    users.forEach(user => {
      let userId = user.id;
      let userEmail = user.email;
      let userGroups = user.groups ? user.groups.join(', ') : ''; // Separa os grupos por vírgula

      // Verificar se os grupos contêm as informações específicas, excluindo "Configuração padrão de Adobe Sign - Corporações TRNS"
      let specificGroups = [
        'Configuração padrão de Todos os Apps - versão Pro - 100 GB',
        'Configuração padrão de Acrobat Pro DC'
      ];

      let groupsToDisplay = user.groups ? user.groups.filter(group => {
        return specificGroups.some(specificGroup => group.includes(specificGroup)) &&
               !group.includes('Configuração padrão de Adobe Sign - Corporações TRNS');
      }) : [];

      // Adicionar "Administrador" se "admin" for encontrado nos grupos
      if (user.groups && user.groups.some(group => group.toLowerCase().includes('admin'))) {
        groupsToDisplay.push('Administrador');
      }

      // Tratar os prefixos e sufixos dos grupos específicos
      groupsToDisplay = groupsToDisplay.map(group => {
        return specificGroups.find(specificGroup => group.includes(specificGroup)) || group;
      });

      // Remover duplicatas, mantendo apenas a informação relevante
      groupsToDisplay = [...new Set(groupsToDisplay.map(group => {
        return group.replace(/_.*$/, ''); // Remove o sufixo
      }))];

      // Se nenhum dos grupos específicos for encontrado, usar o conteúdo completo de groups ou deixar em branco
      if (groupsToDisplay.length === 0) {
        groupsToDisplay = user.groups && user.groups.some(group => group.includes('Configuração padrão de Adobe Sign - Corporações TRNS')) ? '' : userGroups;
      } else {
        groupsToDisplay = groupsToDisplay.join(', ');
      }

      // Substituir qualquer grupo com "_admin" por "Administrador"
      if (user.groups && user.groups.some(group => group.toLowerCase().includes('admin'))) {
        groupsToDisplay = "Administrador";
      }

      sheet.appendRow([userId, userEmail, groupsToDisplay]);
    });
  }
}
