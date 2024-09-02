const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { valuesArray } = require('./../ICDCodes.js'); // Импортируем массив значений

async function checkICDCodes(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];
  const validCodes = new Set(valuesArray); // Используем Set для быстрого поиска

  // Функция для проверки, что код валидный
  function isValidCode(code) {
    return validCodes.has(code.trim());
  }

  // Проверка листа на соответствие кодов
  function checkSheet(sheetName) {
    if (!sheetNames.includes(sheetName)) {
      result.push(`Лист ${sheetName} отсутствует.`);
      return; // Пропустить, если листа нет
    }

    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    for (let i = 1; i < data.length; i++) { // Пропускаем первую строку с заголовками
      const row = data[i];
      const codesCell = row[1]; // Столбец B (индекс 2)

      if (codesCell) {
        // Разбиваем коды на отдельные значения и проверяем их
        const codes = codesCell.toString().split(',').map(code => code.trim());
        const invalidCodes = codes.filter(code => !isValidCode(code));

        if (invalidCodes.length > 0) {
          result.push(`Лист ${sheetName}, строка ${i + 1} содержит невалидные коды: ${invalidCodes.join(', ')}`);
        }
      }
    }
  }

  // Проверяем лист "Диагностика"
  checkSheet('Диагностика');

  // Обработка результатов
  const resultMessage = result.length === 0 ?
    'Все коды в столбце 2 корректны.' :
    result.join('\n');

  const filePath = path.join(__dirname, '..', 'Result.txt');
  fs.writeFileSync(filePath, resultMessage);

  await ctx.replyWithDocument({ source: filePath });
  fs.unlinkSync(filePath);
}

module.exports = { checkICDCodes };