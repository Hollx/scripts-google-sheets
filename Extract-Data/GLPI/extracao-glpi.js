function main_glpi() {
  var sessionToken //= '######################'; teste
  var appToken = '##################'; //toke
  var baseUrl = 'https://###########/glpi/apirest.php'; //link 

  function getSessionToken() {
    var url = baseUrl + '/initSession/';
    var headers = {
      'Authorization': 'user_token #################', //user token glpi
      'App-Token': appToken,
      'Content-Type': 'application/json'
    };

    var response = UrlFetchApp.fetch(url, {
      headers: headers,
      method: 'get'
    });

    if (response.getResponseCode() == 200) {
      var data = JSON.parse(response.getContentText());
      return data['session_token'];
    } else {
      console.error('Error: API request failed with status code ' + response.getResponseCode());
      return null;
    }
  }

  sessionToken = getSessionToken();

  function fetchAllItems(baseUrl, appToken, sessionToken) {
    const items = [];
    let rangeStart = 0;
    const pageSize = 50;

    while (true) {
      const url = `${baseUrl}/Computer?range=${rangeStart}-${rangeStart + pageSize - 1}`;
      const headers = {
        'Session-Token': sessionToken,
        'App-Token': appToken,
        'Content-Type': 'application/json'
      };

      const response = UrlFetchApp.fetch(url, {
        headers: headers,
        method: 'get',
        muteHttpExceptions: true
      });

      const data = JSON.parse(response.getContentText());

      if (response.getResponseCode() == 400 && data[0] === 'ERROR_RANGE_EXCEED_TOTAL') {
        break;
      }

      const extractedItems = data.map((item) => {
        return {
          id: item.id,
          name: item.name,
          contact: item.contact,
          contact_num: item.contact_num, 
          computertypes_id: item.computertypes_id,
          states_id: item.states_id,
          groups_id: item.groups_id,
          locations_id: item.locations_id,
          autoupdatesystems_id: item.autoupdatesystems_id
        };
      });

      items.push(...extractedItems);
      rangeStart += pageSize;
    }

    return items;
  }

  function getDetailsById(endpoint, id) {
    if (!id) return 'N/A'; // Verificar se o ID é válido

    const url = `${baseUrl}/${endpoint}/${id}`;

    const headers = {
      'Session-Token': sessionToken,
      'App-Token': appToken,
      'Content-Type': 'application/json'
    };

    const response = UrlFetchApp.fetch(url, {
      headers: headers,
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() == 200) {
      try {
        const data = JSON.parse(response.getContentText());
        return data.designation || data.name || data.completename || 'N/A'
      } catch (e) {
        Logger.log(`Error parsing JSON for ${endpoint} ID ${id}: ${e}`);
        return 'N/A';
      }
    } else {
      Logger.log(`Failed to get details from ${endpoint} for ID ${id}: ${response.getContentText()}`);
      return 'N/A';
    }
  }

  function getProcessorDetails(computerId) {
    let deviceProcessorId = 'N/A';

    const url = `${baseUrl}/Computer/${computerId}/Item_DeviceProcessor/`;
    const headers = {
      'Session-Token': sessionToken,
      'App-Token': appToken,
      'Content-Type': 'application/json'
    };

    try {
      const response = UrlFetchApp.fetch(url, {
        headers: headers,
        method: 'get',
        muteHttpExceptions: true
      });

      if (response.getResponseCode() == 200) {
        const data = JSON.parse(response.getContentText());

        if (data.length > 0) {
          const deviceProcessorId = data[0].deviceprocessors_id;
          return deviceProcessorId
        }
      } else {
        Logger.log(`Failed to get processor details for Computer ID ${computerId}: ${response.getContentText()}`);
      }
    } catch (error) {
      Logger.log(`Error while fetching processor details for Computer ID ${computerId}: ${error}`);
    }

    return deviceProcessorId;
  }

  function getMemoryDetails(computerId) {
    let memoryDetails = 'N/A';

    const url = `${baseUrl}/Computer/${computerId}/Item_DeviceMemory/`;
    console.log(url)
    const headers = {
      'Session-Token': sessionToken,
      'App-Token': appToken,
      'Content-Type': 'application/json'
    };

    try {
      const response = UrlFetchApp.fetch(url, {
        headers: headers,
        method: 'get',
        muteHttpExceptions: true
      });

      if (response.getResponseCode() == 200) {
        const data = JSON.parse(response.getContentText());
        console.log(data)

        if (data.length > 0) {
          const totalSizeMB = data.reduce((total, item) => total + item.size, 0);
          const totalSizeGB = (totalSizeMB / 1024).toFixed(2); // Converte MB para GB e formata com 2 casas decimais
          return memoryDetails = `${totalSizeGB} GB`;
        }
      } else {
        Logger.log(`Failed to get memory details for Computer ID ${computerId}: ${response.getContentText()}`);
      }
    } catch (error) {
      Logger.log(`Error while fetching memory details for Computer ID ${computerId}: ${error}`);
    }

    return memoryDetails;
  }
  
  function getStorageDetails(computerId) {
    let storageDetails = 'N/A';

    const url = `${baseUrl}/Computer/${computerId}/Item_DeviceHardDrive/`;
    console.log(url)
    const headers = {
      'Session-Token': sessionToken,
      'App-Token': appToken,
      'Content-Type': 'application/json'
    };

    try {
      const response = UrlFetchApp.fetch(url, {
        headers: headers,
        method: 'get',
        muteHttpExceptions: true
      });

      if (response.getResponseCode() == 200) {
        const data = JSON.parse(response.getContentText());
        console.log(data)

        if (data.length > 0) {
          const totalSizeMB = data.reduce((total, item) => total + item.capacity, 0);
          const totalSizeGB = (totalSizeMB / 1024).toFixed(0); // Converte MB para GB e formata com 2 casas decimais
          return storageDetails = `${totalSizeGB} GB`;
        }
      } else {
        Logger.log(`Failed to get memory details for Computer ID ${computerId}: ${response.getContentText()}`);
      }
    } catch (error) {
      Logger.log(`Error while fetching storage details for Computer ID ${computerId}: ${error}`);
    }

    return storageDetails;
  }

  function getAutoUpdateSystemName(id) {
    if (!id) return 'N/A'; // Verificar se o ID é válido

    const url = `${baseUrl}/AutoUpdateSystem/${id}`;
    const headers = {
      'Session-Token': sessionToken,
      'App-Token': appToken,
      'Content-Type': 'application/json'
    };

    const response = UrlFetchApp.fetch(url, {
      headers: headers,
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() == 200) {
      try {
        const data = JSON.parse(response.getContentText());
        return data.name || 'N/A';
      } catch (e) {
        Logger.log(`Error parsing JSON for AutoUpdateSystem ID ${id}: ${e}`);
        return 'N/A';
      }
    } else {
      Logger.log(`Failed to get details for AutoUpdateSystem ID ${id}: ${response.getContentText()}`);
      return 'N/A';
    }
  }

function importToSheet(data) {
    var ss = SpreadsheetApp.openById('###########');
    var sheet = ss.getSheetByName('#########'); // Seleciona a aba 
    sheet.clear();

    const headers = ['ID', 'HostName', 'Usuário', 'Contato', 'Tipo', 'Status', 'Customer', 'Localização', 'Processador', 'RAM', 'Armazenamento', 'Ano'];
    sheet.appendRow(headers);

    data.forEach(item => {
        const computerType = getDetailsById('ComputerType', item.computertypes_id);
        const state = getDetailsById('State', item.states_id);
        const group = getDetailsById('Group', item.groups_id);
        const location = getDetailsById('Location', item.locations_id);
        const processor = getDetailsById("DeviceProcessor", getProcessorDetails(item.id));
        const memory = getMemoryDetails(item.id);
        const storage = getStorageDetails(item.id);
        const autoUpdateSystemName = getAutoUpdateSystemName(item.autoupdatesystems_id);

        sheet.appendRow([item.id, item.name, item.contact, item.contact_num, computerType, state, group, location, processor, memory, storage, autoUpdateSystemName]);
    });
}


  const allComputers = fetchAllItems(baseUrl, appToken, sessionToken);
  importToSheet(allComputers);
}
