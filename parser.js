function parseBalanceData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 4);
  const data = dataRange.getDisplayValues();
  const headers = data[0];
  let config = {};

  for (let col = 1; col < headers.length; col++) {
    const country = headers[col];
    const countryBalance = [];
    let currentBalances = {};

    for (let row = 1; row < data.length; row++) {
      const date = data[row][0];
      const balanceTypes = data[row][col];

      // Если текущих балансов нет, закрываем все текущие балансы
      if (!balanceTypes) {
        for (let balance in currentBalances) {
          if (currentBalances[balance].end_date === null) {
            currentBalances[balance].end_date = date;
          }
          countryBalance.push(currentBalances[balance]);
        }
        currentBalances = {};
        continue;
      }

      const balances = balanceTypes.split(',').map(b => b.trim());

      // Закрываем балансы, которых больше нет
      for (let balance in currentBalances) {
        if (!balances.includes(balance)) {
          if (currentBalances[balance].end_date === null) {
            currentBalances[balance].end_date = date;
          }
          countryBalance.push(currentBalances[balance]);
          delete currentBalances[balance];
        }
      }

      // Добавляем новые балансы и обновляем существующие
      balances.forEach(balance => {
        if (!currentBalances[balance]) {
          currentBalances[balance] = {
            start_date: date,
            end_date: null,
            balance: balance
          };
        } else {
          // Обновляем дату окончания текущего баланса
          currentBalances[balance].end_date = date;
        }
      });
    }

    // Закрываем оставшиеся балансы
    for (let balance in currentBalances) {
      if (currentBalances[balance].end_date === null) {
        currentBalances[balance].end_date = date;
      }
      countryBalance.push(currentBalances[balance]);
    }

    config[`${country}_Balance`] = countryBalance;
  }

  const jsonString = JSON.stringify(config, null, 2);
  Logger.log(jsonString);
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Output') || 
                      SpreadsheetApp.getActiveSpreadsheet().insertSheet('Output');
  outputSheet.getRange(1, 1).setValue(jsonString);
}
