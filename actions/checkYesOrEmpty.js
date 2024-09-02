const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function checkYesOrEmpty(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];

  // Функция для проверки, что ячейка пустая или содержит "да"
  function isValidCellValue(cell) {
    const value = cell ? cell.toString().trim().toLowerCase() : '';
    return value === '' || value === 'да';
  }

  // Проверка листа на правильность значений в определенных столбцах
  function checkSheet(sheetName, columns) {
    if (!sheetNames.includes(sheetName)) {
      result.push(`Лист ${sheetName} отсутствует.`);
      return; // Пропустить, если листа нет
    }

    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    for (let i = 1; i < data.length; i++) { // Пропускаем первую строку с заголовками
      const row = data[i];
      let invalidCells = [];

      columns.forEach(colIdx => {
        const cell = row[colIdx];
        if (!isValidCellValue(cell)) {
          const columnName = String.fromCharCode(65 + colIdx); // Преобразуем индекс в букву столбца
          invalidCells.push({ column: columnName, value: cell });
        }
      });

      if (invalidCells.length > 0) {
        const invalidDetails = invalidCells.map(cell => `столбец ${cell.column}: "${cell.value}"`).join(', ');
        result.push(`Лист ${sheetName}, строка ${i + 1} содержит некорректные значения: ${invalidDetails}`);
      }
    }
  }

  // Проверяем лист "Диагностика"
  checkSheet('Диагностика', [17, 16, 10, 6]); // Столбцы R (17), Q (16), K (10), G (6)

  // Обработка результатов
  const resultMessage = result.length === 0 ?
    'Все значения в указанных столбцах корректны.' :
    result.join('\n');

  const filePath = path.join(__dirname, '..', 'Result.txt');
  fs.writeFileSync(filePath, resultMessage);

  await ctx.replyWithDocument({ source: filePath });
  fs.unlinkSync(filePath);
}

module.exports = { checkYesOrEmpty };