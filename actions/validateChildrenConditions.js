const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function validateChildrenConditions(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];

  if (sheetNames.length === 0) {
    result.push('Файл не содержит листов.');
  } else {
    const sheetName = sheetNames[0]; // Первый лист
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    for (let i = 1; i < data.length; i++) { // Пропускаем заголовки
      const row = data[i];
      const valueA = row[0]; // Столбец A (идентификатор)
      const valueB = row[1]?.toString(); // Столбец B (строка)
      const valueM = parseFloat(row[12]); // Столбец M (индекс 12)
      const valueN = parseFloat(row[13]); // Столбец N (индекс 13)

      // Пропускаем строки, если значение M не является числом
      if (isNaN(valueM)) continue;

      // Условия проверки
      if (/дети|детей/i.test(valueB)) {
        // Если в столбце B есть "дети" или "детей"
        if (valueM > 214) {
          result.push(
            `Лист "${sheetName}", строка ${i + 1}: ИД="${valueA}", ошибка в столбце M, значение "${valueM}"`
          );
        }
        if (!isNaN(valueN) && valueN > 215) {
          result.push(
            `Лист "${sheetName}", строка ${i + 1}: ИД="${valueA}", ошибка в столбце N, значение "${valueN}"`
          );
        }
      } else {
        // Если в столбце B нет "дети" или "детей"
        if (valueM < 215) {
          result.push(
            `Лист "${sheetName}", строка ${i + 1}: ИД="${valueA}", ошибка в столбце M, значение "${valueM}"`
          );
        }
      }
    }
  }

  // Обработка результатов
  const resultMessage =
    result.length === 0
      ? 'Все строки прошли валидацию на первом листе.'
      : result.join('\n');

  const filePath = path.join(__dirname, '..', 'ValidationResult.txt');
  fs.writeFileSync(filePath, resultMessage);

  await ctx.replyWithDocument({ source: filePath });
  fs.unlinkSync(filePath);
}

module.exports = { validateChildrenConditions };