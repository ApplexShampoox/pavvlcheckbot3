const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function validateUniqueIDs(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  const idMap = new Map(); // Хранит все ИД и листы, где они встречаются
  let result = [];

  // Сканируем все листы
  sheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    for (let i = 1; i < data.length; i++) { // Пропускаем заголовки
      const row = data[i];
      const valueID = row[0]?.toString().trim(); // Первый столбец (ИД)

      if (!valueID) continue; // Пропуск пустых строк

      if (!idMap.has(valueID)) {
        idMap.set(valueID, new Set());
      }

      idMap.get(valueID).add(sheetName); // Добавляем лист в список для этого ИД
    }
  });

  // Проверяем, на всех ли листах встречаются ИД
  idMap.forEach((sheets, id) => {
    if (sheets.size !== sheetNames.length) {
      const missingSheets = sheetNames.filter((sheetName) => !sheets.has(sheetName));
      result.push(
        `ИД "${id}" отсутствует на листах: ${missingSheets.join(', ')}`
      );
    }
  });

  // Обработка результатов
  const resultMessage =
    result.length === 0
      ? 'Все ИД из первого столбца присутствуют на всех листах.'
      : result.join('\n');

  const filePath = path.join(__dirname, '..', 'IDValidationResult.txt');
  fs.writeFileSync(filePath, resultMessage);

  await ctx.replyWithDocument({ source: filePath });
  fs.unlinkSync(filePath);
}

module.exports = { validateUniqueIDs };