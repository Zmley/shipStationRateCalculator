function postShippingRates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const batchSize = 10; // 每次处理的行数
  const markerColumn = 18; // 标记所在的列号

  // 找到标记行号
  let startRow = findStartRow(sheet, markerColumn, lastRow);

  // 如果没有标记，默认从第2行开始
  if (!startRow) {
    startRow = 2;
  }

  const endRow = Math.min(startRow + batchSize - 1, lastRow);

  const username = "";
  const password = "";

  const targetServices = [
    { serviceCode: "fedex_ground", serviceName: "FedEx Ground" },
    { serviceCode: "fedex_home_delivery", serviceName: "FedEx Home Delivery" },
    { serviceCode: "ups_ground", serviceName: "UPS Ground" },
    { serviceCode: "usps_parcel_select", serviceName: "USPS Parcel Select Ground" },
    { serviceCode: "usps_ground_advantage", serviceName: "USPS Ground Advantage" }
  ];

  for (let i = startRow; i <= endRow; i++) {
    const fromPostalCode = sheet.getRange(i, 2).getValue();
    const toState = sheet.getRange(i, 3).getValue();
    const toCountry = sheet.getRange(i, 4).getValue();
    const toPostalCode1 = sheet.getRange(i, 5).getValue();
    const toPostalCode2 = sheet.getRange(i, 6).getValue();
    const toCity = sheet.getRange(i, 7).getValue();
    const weight = sheet.getRange(i, 8).getValue();
    const length = sheet.getRange(i, 9).getValue();
    const width = sheet.getRange(i, 10).getValue();
    const height = sheet.getRange(i, 11).getValue();
    const residential = true; // 强制为 true

    let allResults = [];

    const toPostalCodes = [
      { toPostalCode: toPostalCode1, label: "to postal code 1" },
      ...(toPostalCode2 ? [{ toPostalCode: toPostalCode2, label: "to postal code 2" }] : [])
    ];

    const carriers = [
      { code: "fedex", name: "FedEx" },
      { code: "ups", name: "UPS" },
      { code: "stamps_com", name: "USPS" }
    ];

    for (let toPostalCodeObj of toPostalCodes) {
      const { toPostalCode, label } = toPostalCodeObj;
      const payload = {
        fromPostalCode,
        toState,
        toCountry: "US",
        toPostalCode,
        toCity,
        weight: { value: weight, units: "pounds" },
        dimensions: { units: "inches", length, width, height },
        confirmation: "delivery",
        residential
      };

      for (let carrier of carriers) {
        try {
          payload.carrierCode = carrier.code;
          const options = {
            method: "post",
            contentType: "application/json",
            headers: { "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password) },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          };

          const response = UrlFetchApp.fetch("https://ssapi.shipstation.com/shipments/getrates", options);
          const responseCode = response.getResponseCode();

          if (responseCode === 500) {
            continue;
          }

          const result = JSON.parse(response.getContentText());

          const filteredResult = result.filter(service =>
            targetServices.some(target => target.serviceCode === service.serviceCode)
          );

          if (filteredResult.length > 0) {
            allResults.push(
              ...filteredResult.map(service => ({
                carrier: carrier.name,
                name: service.serviceName,
                price: service.shipmentCost + service.otherCost,
                postalCodeLabel: label
              }))
            );
          }
        } catch (error) {
          console.log(`Error with ${carrier.name} (${label}): ${error.message}`);
        }
      }
    }

    // 写入结果到表格
    writeResults(sheet, i, allResults, toPostalCode2);
  }

  // 清除当前标记并设置新的标记
  sheet.getRange(startRow, markerColumn).clearContent();
  if (endRow < lastRow) {
    sheet.getRange(endRow + 1, markerColumn).setValue("START");
    createNextTrigger(); // 创建新的触发器
  } else {
    SpreadsheetApp.getUi().alert("所有数据已成功发送并处理完成！");
    clearAllTriggers(); // 删除所有触发器
  }
}

function findStartRow(sheet, markerColumn, lastRow) {
  for (let i = 2; i <= lastRow; i++) {
    const value = sheet.getRange(i, markerColumn).getValue();
    if (value === "START") {
      return i;
    }
  }
  return null;
}

function writeResults(sheet, rowIndex, allResults, toPostalCode2) {
  if (allResults.length > 0) {
    const allResultsText = allResults
      .map(result => `${result.carrier}: $${result.price.toFixed(2)} (${result.name}, ${result.postalCodeLabel})`)
      .join("\n");
    sheet.getRange(rowIndex, 13).setValue(allResultsText);
  } else {
    sheet.getRange(rowIndex, 13).setValue("No rates available");
  }

  const resultsToPostal1 = allResults.filter(result => result.postalCodeLabel === "to postal code 1");
  const resultsToPostal2 = allResults.filter(result => result.postalCodeLabel === "to postal code 2");

  writeLowestPrice(sheet, rowIndex, 14, 15, resultsToPostal1);
  if (toPostalCode2) {
    writeLowestPrice(sheet, rowIndex, 16, 17, resultsToPostal2);
  }
}

function writeLowestPrice(sheet, rowIndex, methodCol, priceCol, results) {
  if (results.length > 0) {
    const minPrice = Math.min(...results.map(result => result.price));
    const lowestResults = results.filter(result => result.price === minPrice);
    const lowestMethod = lowestResults
      .map(result => `${result.carrier} (${result.name})`)
      .join("\n");
    sheet.getRange(rowIndex, methodCol).setValue(lowestMethod);
    sheet.getRange(rowIndex, priceCol).setValue(`$${minPrice.toFixed(2)}`);
  } else {
    sheet.getRange(rowIndex, methodCol).setValue("No methods available");
    sheet.getRange(rowIndex, priceCol).setValue("No price available");
  }
}

function createNextTrigger() {
  clearAllTriggers(); // 先删除旧的触发器
  ScriptApp.newTrigger("postShippingRates")
    .timeBased()
    .after(5000) // 5 秒后运行
    .create();
}

function clearAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}