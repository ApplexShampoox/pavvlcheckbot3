const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

async function checkRequiredColumns(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];

  // Функция для проверки, что ячейка пуста
  function isCellEmpty(cell) {
    return cell === undefined || cell === null || cell.toString().trim() === '';
  }

  // Функция для проверки, что строка не полностью пустая
  function isRowNotEmpty(row) {
    return row.some(cell => !isCellEmpty(cell));
  }

  // Функция для получения буквенного представления столбца по его индексу
  function getColumnLetter(index) {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }

  // Проверка листа на наличие незаполненных обязательных столбцов
  function checkSheet(sheetName, requiredColumns) {
    if (!sheetNames.includes(sheetName)) return; // Пропустить, если листа нет

    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Получаем заголовки столбцов из первой строки
    const headers = data[0];

    for (let i = 1; i < data.length; i++) { // Пропускаем первую строку с заголовками
      const row = data[i];

      // Пропуск пустых строк
      if (!isRowNotEmpty(row)) continue;

      let missingColumns = [];

      requiredColumns.forEach(colIdx => {
        const cell = row[colIdx];
        if (isCellEmpty(cell)) {
          const columnName = headers[colIdx]; // Используем заголовок столбца
          const columnLetter = getColumnLetter(colIdx); // Получаем букву столбца
          missingColumns.push(`${columnName} (${columnLetter})`);
        }
      });

      // Если в строке есть незаполненные обязательные столбцы, добавляем их в результат
      if (missingColumns.length > 0) {
        result.push(`Лист ${sheetName}, строка ${i + 1} не заполнены столбцы: ${missingColumns.join(', ')}`);
      }
    }
  }

  // Проверяем лист "Диагностика"
  checkSheet('Диагностика', [0, 1, 2, 3, 4, 5, 7, 8, 13, 15]); // Столбцы A, B, C, D, E, F, H, I, N, P

  // Проверяем лист "Лечение"
  checkSheet('Лечение', [0, 1, 2, 3, 5, 25, 27]); // Столбцы A, B, C, D, F, Z, AB

  // Проверяем лист "Профильный специалист"
  checkSheet('Профильный специалист', [0, 1, 2, 3]); // Столбцы A, B, C, D

  // Проверяем лист "Шаблоны"
  checkSheet('Шаблоны', [0, 1, 2, 3, 4, 5]); // Столбцы A, B, C, D, E, F

  // Обработка результатов
  if (result.length === 0) {
    await ctx.reply('Все обязательные столбцы заполнены.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });

    fs.unlinkSync(filePath);
  }
}

module.exports = { checkRequiredColumns };