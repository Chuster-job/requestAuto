// Глобальные константы (должны быть заданы в свойствах скрипта или другом месте)
const CLIENT_ID = clientId; // Замените на реальный Client-ID
const API_KEY = apiKey;     // Замените на реальный API-ключ

// Основная функция для получения и обработки данных
function fetchDataAndProcess() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист31");
  
  // Очистка только необходимых столбцов (A, B, E, F, G)
  clearSheetColumns(sheet);
  
  // Запись заголовков
  setHeaders(sheet);
  
  // Таблица соответствий регионов и их ID
  const regionToClusterId = getRegionClusterMapping();
  
  try {
    Logger.log("Начинаем запрос к API Ozon для получения списка кластеров.");
    
    const regionName = sheet.getRange("H1").getValue();
    Logger.log(`Получено название региона из ячейки H1: ${regionName}`);
    
    const clusterId = getClusterId(regionName, regionToClusterId);
    Logger.log(`ID кластера для региона "${regionName}": ${clusterId}`);
    
    // Получаем данные кластеров
    const clusterData = fetchClusterData(clusterId);
    Logger.log(`Получено ${clusterData.clusters?.length || 0} кластеров.`);
    
    if (clusterData.clusters && clusterData.clusters.length > 0) {
      processClusters(sheet, clusterData.clusters);
    } else {
      Logger.log("No clusters found in the response.");
    }
  } catch (error) {
    Logger.log(`Критическая ошибка: ${error.message}`);
    throw error; // Пробрасываем ошибку дальше для обработки
  }
}

// Вспомогательные функции

function clearSheetColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent(); // Столбец A (Cluster ID)
    sheet.getRange(2, 2, lastRow - 1, 1).clearContent(); // Столбец B (Cluster Name)
    sheet.getRange(2, 5, lastRow - 1, 1).clearContent(); // Столбец E (Operation ID)
    sheet.getRange(2, 6, lastRow - 1, 1).clearContent(); // Столбец F (Warehouse Name)
    sheet.getRange(2, 7, lastRow - 1, 1).clearContent(); // Столбец G (Warehouse ID)
  }
}

function setHeaders(sheet) {
  sheet.getRange(1, 1).setValue("Cluster ID");
  sheet.getRange(1, 2).setValue("Cluster Name");
  sheet.getRange(1, 5).setValue("Operation ID");
  sheet.getRange(1, 6).setValue("Warehouse Name");
  sheet.getRange(1, 7).setValue("Warehouse ID");
}

function getRegionClusterMapping() {
  return {
    "Санкт-Петербург и СЗО": 2,
    "Урал": 3,
    "Дальний Восток": 7,
    "Калининград": 12,
    "Воронеж": 16,
    "Краснодар": 17,
    "Тюмень": 144,
    "Волгоград": 146,
    "Ростов": 147,
    "Уфа": 148,
    "Казань": 149,
    "Самара": 150,
    "Новосибирск": 151,
    "Омск": 152,
    "Кавказ": 153,
    "Москва, МО и Дальние регионы": 154,
    "Красноярск": 155
  };
}

function getClusterId(regionName, mapping) {
  const clusterId = mapping[regionName];
  if (!clusterId) {
    throw new Error(`Не найден ID кластера для региона: ${regionName}`);
  }
  return clusterId;
}

function fetchClusterData(clusterId) {
  const response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/cluster/list", {
    method: "post",
    contentType: "application/json",
    headers: {
      "Client-Id": CLIENT_ID,
      "Api-Key": API_KEY
    },
    payload: JSON.stringify({
      "cluster_ids": [clusterId.toString()],
      "cluster_type": "CLUSTER_TYPE_OZON"
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    const errorDetails = response.getContentText();
    Logger.log(`Ошибка при получении списка кластеров. Код: ${response.getResponseCode()}, Ответ: ${errorDetails}`);
    throw new Error(`Ошибка при получении списка кластеров: ${errorDetails}`);
  }

  return JSON.parse(response.getContentText());
}

function processClusters(sheet, clusters) {
  clusters.forEach((cluster, index) => {
    const row = index + 2; // Начинаем с 2 строки
    sheet.getRange(row, 1).setValue(cluster.id);
    sheet.getRange(row, 2).setValue(cluster.name);
    Logger.log(`Обрабатываем кластер с ID: ${cluster.id}, Name: ${cluster.name}`);

    try {
      const sku = sheet.getRange(row, 3).getValue(); // Столбец C
      const quantity = sheet.getRange(row, 4).getValue(); // Столбец D

      validateInputs(sku, quantity, row, sheet);

      const operationId = sendDraftRequestWithRetry(cluster.id, sku, quantity);
      sheet.getRange(row, 5).setValue(operationId);

      const infoData = checkCalculationStatus(operationId);
      
      if (infoData.clusters && infoData.clusters.length > 0) {
        showWarehouseSidebar(infoData, row, operationId);
      } else {
        Logger.log(`Для operation_id: ${operationId} не найдено данных о кластерах.`);
        sheet.getRange(row, 6).setValue("Кластеры не найдены");
        sheet.getRange(row, 7).setValue("");
      }
    } catch (error) {
      const errorMessage = `Ошибка при обработке кластера ID: ${cluster.id}: ${error.message}`;
      Logger.log(errorMessage);
      sheet.getRange(row, 5).setValue(errorMessage);
    }
  });
}
// Функция для отображения Sidebar с выпадающим списком складов
function showWarehouseSidebar(infoData, row, operationId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист31");

  // Извлекаем данные о складах из infoData
  const clusterInfo = infoData.clusters && infoData.clusters.length > 0 ? infoData.clusters[0] : null;
  if (!clusterInfo || !clusterInfo.warehouses || clusterInfo.warehouses.length === 0) {
    Logger.log("Нет данных о складах для отображения.");
    return;
  }

  // Формируем массив данных о складах
  const warehouses = clusterInfo.warehouses.map(warehouse => {
    const supplyWarehouse = warehouse.supply_warehouse || {};
    const name = supplyWarehouse.name || "Название не указано";
    const id = supplyWarehouse.warehouse_id || "ID не указано";
    return { name, id };
  });

  Logger.log(`Сформированные данные о складах: ${JSON.stringify(warehouses)}`);

  // Формируем HTML-код для выпадающего списка
  const warehousesHtml = warehouses.map(warehouse =>
    `<option value="${warehouse.id}">${warehouse.name} (ID: ${warehouse.id})</option>`
  ).join("");

  // Отображаем Sidebar
  const template = HtmlService.createTemplateFromFile("sidebar");
  template.warehousesHtml = warehousesHtml;
  const htmlOutput = template.evaluate().setTitle("Выбор склада");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);

  // Сохраняем данные о строке и operation_id для последующего использования
  PropertiesService.getScriptProperties().setProperty("currentRow", row.toString());
  PropertiesService.getScriptProperties().setProperty("operationId", operationId.toString());
}


function validateInputs(sku, quantity, row, sheet) {
  if (!sku || !Number.isInteger(sku)) {
    const errorMessage = `Ошибка: Некорректное значение SKU для строки ${row}`;
    Logger.log(errorMessage);
    sheet.getRange(row, 5).setValue(errorMessage);
    throw new Error(errorMessage);
  }
  
  if (!quantity || !Number.isInteger(quantity)) {
    const errorMessage = `Ошибка: Некорректное значение Quantity для строки ${row}`;
    Logger.log(errorMessage);
    sheet.getRange(row, 5).setValue(errorMessage);
    throw new Error(errorMessage);
  }
}

function sendDraftRequestWithRetry(clusterId, sku, quantity, retries = 3, delay = 2000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      Logger.log(`Попытка ${attempt} для кластера ID: ${clusterId}, SKU: ${sku}, Quantity: ${quantity}`);
      
      const draftResponse = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/draft/create", {
        method: "post",
        contentType: "application/json",
        headers: {
          "Client-Id": CLIENT_ID,
          "Api-Key": API_KEY
        },
        payload: JSON.stringify({
          "cluster_ids": [clusterId.toString()],
          "drop_off_point_warehouse_id": 21957475354000,
          "items": [
            {
              "quantity": quantity,
              "sku": sku
            }
          ],
          "type": "CREATE_TYPE_CROSSDOCK"
        }),
        muteHttpExceptions: true
      });

      if (draftResponse.getResponseCode() === 200) {
        const draftData = JSON.parse(draftResponse.getContentText());
        Logger.log(`Успешно создан черновик для кластера ID: ${clusterId}. Operation ID: ${draftData.operation_id}`);
        saveOperationId(draftData.operation_id);
        return draftData.operation_id;
      } else if (draftResponse.getResponseCode() === 429) {
        Logger.log(`Ошибка 429 для кластера ID: ${clusterId}. Ждем ${delay} мс перед повторной попыткой.`);
        Utilities.sleep(delay);
      } else {
        const errorDetails = draftResponse.getContentText();
        throw new Error(`Ошибка: ${errorDetails}`);
      }
    } catch (error) {
      Logger.log(`Ошибка при создании черновика для кластера ID: ${clusterId}: ${error.message}`);
      if (attempt === retries) throw error;
    }
  }
  throw new Error(`Превышено количество попыток (${retries}) для кластера ID: ${clusterId}`);
}

function saveOperationId(operationId) {
  PropertiesService.getScriptProperties().setProperty("operationId", operationId.toString());
  Logger.log(`operationId сохранен: ${operationId}`);
}

function checkCalculationStatus(operationId, retries = 5, delay = 5000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      Logger.log(`Попытка ${attempt} для проверки статуса расчета с operation_id: ${operationId}`);
      
      const infoResponse = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/draft/create/info", {
        method: "post",
        contentType: "application/json",
        headers: {
          "Client-Id": CLIENT_ID,
          "Api-Key": API_KEY
        },
        payload: JSON.stringify({
          "operation_id": operationId
        }),
        muteHttpExceptions: true
      });

      if (infoResponse.getResponseCode() !== 200) {
        const errorDetails = infoResponse.getContentText();
        throw new Error(`Ошибка при запросе к /v1/draft/create/info: ${errorDetails}`);
      }

      const infoData = JSON.parse(infoResponse.getContentText());
      Logger.log(`Получен ответ от /v1/draft/create/info: ${JSON.stringify(infoData)}`);

      switch (infoData.status) {
        case "CALCULATION_STATUS_IN_PROGRESS":
          Logger.log(`Статус расчета: IN_PROGRESS. Ждем ${delay} мс перед повторной попыткой.`);
          Utilities.sleep(delay);
          break;
        case "CALCULATION_STATUS_FAILED":
          handleCalculationFailed(infoData);
          break;
        case "CALCULATION_STATUS_SUCCESS":
          Logger.log(`Статус расчета: SUCCESS.`);
          return infoData;
        default:
          throw new Error(`Неизвестный статус расчета: ${infoData.status}`);
      }
    } catch (error) {
      Logger.log(`Ошибка при проверке статуса расчета: ${error.message}`);
      if (attempt === retries) throw error;
    }
  }
  throw new Error(`Превышено количество попыток проверки статуса расчета для operation_id: ${operationId}`);
}

function handleCalculationFailed(infoData) {
  Logger.log(`Статус расчета: FAILED. Обрабатываем ошибки.`);
  let errorMessage = "Ошибки: ";
  
  if (infoData.errors && infoData.errors.length > 0) {
    infoData.errors.forEach(error => {
      if (error.error_message) {
        errorMessage += `${error.error_message}. `;
      }
      if (error.items_validation && error.items_validation.length > 0) {
        error.items_validation.forEach(itemValidation => {
          errorMessage += `SKU ${itemValidation.sku}: ${itemValidation.reasons.join(", ")}. `;
        });
      }
    });
  }
  
  throw new Error(errorMessage);
}


// Обновленная функция processSelectedWarehouseId
function processSelectedWarehouseId(selectedId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист31");
  const currentRow = parseInt(PropertiesService.getScriptProperties().getProperty("currentRow"), 10);
  const operationId = PropertiesService.getScriptProperties().getProperty("operationId");
  
  if (!currentRow || !operationId) {
    throw new Error("Не найдены currentRow или operationId");
  }

  try {
    // 1. Получаем актуальную информацию о черновике
    const draftInfo = getDraftInfoWithRetry(operationId);
    
    // 2. Проверяем существование черновика
    if (!draftInfo.draft_id) {
      throw new Error("Черновик не найден или был удален");
    }
    
    // 3. Получаем таймслоты для выбранного склада
    const timeslots = fetchTimeslotWithRetry(draftInfo.draft_id, selectedId);
    
    // 4. Отображаем диалог выбора таймслота
    showTimeslotSelectionDialog(timeslots, currentRow, selectedId);
    
  } catch (error) {
    Logger.log(`Ошибка: ${error.message}`);
    sheet.getRange(currentRow, 7).setValue(`Ошибка: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Ошибка: ${error.message}`);
  }
}


// Обновлённая функция получения таймслотов
function fetchTimeslotWithRetry(draftId, warehouseId, retries = 3, delay = 3000) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист31");
  const today = new Date();
  const dateFrom = formatDate(today);
  const dateTo = formatDate(new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000));

  const payload = {
    "date_from": dateFrom,
    "date_to": dateTo,
    "draft_id": draftId,
    "warehouse_ids": [warehouseId]
  };

  try {
    const response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/draft/timeslot/info", {
      method: "post",
      headers: {
        "Client-Id": CLIENT_ID,
        "Api-Key": API_KEY,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const rawResponse = response.getContentText();
    sheet.getRange("I1").setValue(rawResponse); // Для отладки

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(rawResponse);
      
      // Извлекаем все таймслоты из сложной структуры
      const allTimeslots = [];
      data.drop_off_warehouse_timeslots?.forEach(warehouse => {
        warehouse.days?.forEach(day => {
          day.timeslots?.forEach(slot => {
            allTimeslots.push({
              from: slot.from_in_timezone,
              to: slot.to_in_timezone,
              warehouse_id: warehouse.drop_off_warehouse_id
            });
          });
        });
      });

      if (allTimeslots.length === 0) {
        throw new Error("Нет доступных таймслотов в ответе API");
      }
      
      return allTimeslots;
    } else {
      throw new Error(`API Error: ${rawResponse}`);
    }
  } catch (error) {
    sheet.getRange("I1").setValue("ОШИБКА: " + error.message);
    throw error;
  }
}

// Новая функция для отображения диалога выбора таймслота
function showTimeslotSelectionDialog(timeslots, row, warehouseId) {
  const htmlTemplate = HtmlService.createTemplateFromFile('timeslotDialog');
  const timezone = Session.getScriptTimeZone();
  
  // Форматируем таймслоты для отображения
  htmlTemplate.timeslots = timeslots.map(slot => {
    const start = new Date(slot.from);
    const end = new Date(slot.to);
    return {
      value: slot.from, // Используем from как уникальный идентификатор
      label: `${Utilities.formatDate(start, timezone, "dd.MM.yyyy HH:mm")} - ${Utilities.formatDate(end, timezone, "HH:mm")}`,
      duration: ((end - start) / (1000 * 60 * 60)).toFixed(1) + " ч."
    };
  });
  
  htmlTemplate.row = row;
  htmlTemplate.warehouseId = warehouseId;
  
  const html = htmlTemplate.evaluate()
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Выбор таймслота');
}
// Функция сохранения выбранного таймслота
function saveSelectedTimeslot(row, warehouseId, slotIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Лист31");
  
  // Сохраняем выбранный склад и таймслот
  sheet.getRange(row, 6).setValue("Склад выбран");
  sheet.getRange(row, 7).setValue(warehouseId);
  
  // Можно сохранить дополнительную информацию о таймслоте
  // Например, в столбцы I, J, K:
  sheet.getRange(row, 9).setValue("Выбранный таймслот: " + slotIndex);
  
  Logger.log(`Для строки ${row} выбран таймслот с индексом ${slotIndex}`);
}
// Новая функция для получения draft_id
function getDraftInfoWithRetry(operationId, retries = 3, delay = 2000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      Logger.log(`Попытка ${attempt}/${retries} получения draft_id для operationId: ${operationId}`);
      
      const response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/draft/create/info", {
        method: "post",
        headers: {
          "Client-Id": CLIENT_ID,
          "Api-Key": API_KEY,
          "Content-Type": "application/json"
        },
        payload: JSON.stringify({ "operation_id": operationId }),
        muteHttpExceptions: true
      });

      const responseData = JSON.parse(response.getContentText());
      
      if (response.getResponseCode() === 200) {
        Logger.log(`Успешно получен draft_id: ${responseData.draft_id}`);
        return responseData;
      } else {
        throw new Error(`API Error: ${responseData.message || response.getContentText()}`);
      }
    } catch (error) {
      if (attempt === retries) throw error;
      Logger.log(`Ошибка (попытка ${attempt}): ${error.message}. Ждем ${delay} мс...`);
      Utilities.sleep(delay);
    }
  }
  throw new Error(`Не удалось получить draft_id после ${retries} попыток`);
}

// Форматирование даты (остается без изменений)
function formatDate(date) {
  const pad = (num) => num.toString().padStart(2, '0');
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}Z`;
}
