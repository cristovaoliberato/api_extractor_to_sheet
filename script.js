function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function sendRequestWithPagination(url, params, pagination, fileName) {
  const results = fetchPaginatedData(url, params, pagination);
  const spreadsheet = createSpreadsheetFromResults(results, fileName || "Resultado API");
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  const finalFileUrl = convertSpreadsheetToXlsx(spreadsheet, fileName);
  return finalFileUrl;
}

function fetchPaginatedData(url, params, pagination) {
  const results = [];
  const pageSize = parseInt(pagination.sizeCount || 10, 10);
  const totalPages = pagination.isPaginated ? parseInt(pagination.pageCount || 1, 10) : 1;

  for (let i = 0; i < totalPages; i++) {
    const { requestUrl, options } = buildRequest(url, params, pagination, i, pageSize);
    const response = UrlFetchApp.fetch(requestUrl, options);
    const content = response.getContentText();

    try {
      const parsed = JSON.parse(content);
      results.push(parsed);
    } catch (e) {
      results.push({ error: "Invalid JSON", raw: content });
    }
  }
  return results;
}

function buildRequest(url, params, pagination, pageIndex, pageSize) {
  const payload = {};
  const options = {
    headers: { ...params }
  };

  if (pagination.controlType === "query") {
    payload[pagination.pageParam] = pageIndex;
    payload[pagination.sizeParam] = pageSize;
    const queryString = new URLSearchParams(payload).toString();
    return {
      requestUrl: `${url}?${queryString}`,
      options
    };
  }

  if (pagination.controlType === "body") {
    payload[pagination.pageParam] = pageIndex;
    payload[pagination.sizeParam] = pageSize;
    return {
      requestUrl: url,
      options: {
        ...options,
        method: "GET",
        contentType: "application/json",
        payload: JSON.stringify(payload)
      }
    };
  }

  if (pagination.controlType === "path") {
    return {
      requestUrl: `${url}?${pagination.pageParam}=${pageIndex}&${pagination.sizeParam}=${pageSize}`,
      options: {
        ...options,
        method: "GET"
      }
    };
  }

  return { requestUrl: url, options };
}

function createSpreadsheetFromResults(results, fileName) {
  const tempSpreadsheet = SpreadsheetApp.create(fileName);
  const sheet = tempSpreadsheet.getActiveSheet();

  const allData = [];

  results.forEach(result => {
    if (Array.isArray(result)) {
      result.forEach(item => {
        if (typeof item === 'object') {
          allData.push(flattenObject(item));
        }
      });
    } else if (typeof result === 'object') {
      allData.push(flattenObject(result));
    }
  });


  if (allData.length > 0) {
    const headers = Object.keys(allData[0]);
    const rows = allData.map(obj => headers.map(h => obj[h] || ""));

    Logger.log("Headers: " + JSON.stringify(headers));
    Logger.log("Primeira linha: " + JSON.stringify(rows[0]));

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  } else {
    Logger.log("Nenhum dado encontrado para salvar na planilha.");
  }

  return tempSpreadsheet;
}

function convertSpreadsheetToXlsx(spreadsheet, fileName) {
  const fileId = spreadsheet.getId();
   const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + token
    }
  });

  const xlsxBlob = response.getBlob().setName(fileName.endsWith(".xlsx") ? fileName : fileName + ".xlsx");
  const finalFile = DriveApp.createFile(xlsxBlob);
  DriveApp.getFileById(fileId).setTrashed(true);
  return finalFile.getUrl();
} 


function flattenObject(obj, prefix = '', res = {}) {
  for (let key in obj) {
    const prop = prefix ? `${prefix}.${key}` : key;
    if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
      flattenObject(obj[key], prop, res);
    } else {
      res[prop] = obj[key];
    }
  }
  return res;
}
