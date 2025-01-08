function postShippingRates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  const username = "";
  const password = "";

  const targetServices = [
    { serviceCode: "fedex_ground", serviceName: "FedEx Ground" },
    { serviceCode: "fedex_home_delivery", serviceName: "FedEx Home Delivery" },
    { serviceCode: "ups_ground", serviceName: "UPS Ground" },
    { serviceCode: "usps_parcel_select", serviceName: "USPS Parcel Select Ground" },
    { serviceCode: "usps_ground_advantage", serviceName: "USPS Ground Advantage" }
  ];

  for (let i = 2; i <= lastRow; i++) {
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
    const residential = String(sheet.getRange(i, 12).getValue()).toLowerCase() === "true";

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
        "fromPostalCode": fromPostalCode,
        "toState": toState,
        "toCountry": "US",
        "toPostalCode": toPostalCode,
        "toCity": toCity,
        "weight": { "value": weight, "units": "pounds" },
        "dimensions": { "units": "inches", "length": length, "width": width, "height": height },
        "confirmation": "delivery",
        "residential": true
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

    // 将所有结果写入 M 列
    if (allResults.length > 0) {
      const allResultsText = allResults
        .map(result => `${result.carrier}: $${result.price.toFixed(2)} (${result.name}, ${result.postalCodeLabel})`)
        .join("\n");
      sheet.getRange(i, 13).setValue(allResultsText); // M 列
    } else {
      sheet.getRange(i, 13).setValue("No rates available");
    }

    // 分别处理最低价格和方法 (Shipping Method)
    const resultsToPostal1 = allResults.filter(result => result.postalCodeLabel === "to postal code 1");
    const resultsToPostal2 = allResults.filter(result => result.postalCodeLabel === "to postal code 2");

    if (resultsToPostal1.length > 0) {
      const minPrice1 = Math.min(...resultsToPostal1.map(result => result.price));
      const lowestResults1 = resultsToPostal1.filter(result => result.price === minPrice1);
      const lowestMethod1 = lowestResults1
        .map(result => `${result.carrier} (${result.name})`)
        .join("\n");
      sheet.getRange(i, 14).setValue(lowestMethod1); // Shipping Method to Code1
      sheet.getRange(i, 15).setValue(`$${minPrice1.toFixed(2)}`); // Lowest Price to Code1
    } else {
      sheet.getRange(i, 14).setValue("No methods available");
      sheet.getRange(i, 15).setValue("No price available");
    }

    if (toPostalCode2 && resultsToPostal2.length > 0) {
      const minPrice2 = Math.min(...resultsToPostal2.map(result => result.price));
      const lowestResults2 = resultsToPostal2.filter(result => result.price === minPrice2);
      const lowestMethod2 = lowestResults2
        .map(result => `${result.carrier} (${result.name})`)
        .join("\n");
      sheet.getRange(i, 16).setValue(lowestMethod2); // Shipping Method to Code2
      sheet.getRange(i, 17).setValue(`$${minPrice2.toFixed(2)}`); // Lowest Price to Code2
    } else if (toPostalCode2) {
      sheet.getRange(i, 16).setValue("No methods available");
      sheet.getRange(i, 17).setValue("No price available");
    }
  }

  SpreadsheetApp.getUi().alert("所有数据已成功发送并处理完成！");
}