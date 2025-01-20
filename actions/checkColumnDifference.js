const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function checkColumnDifference(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];

  // Проверяем первые два листа
  for (let sheetIndex = 0; sheetIndex < 2; sheetIndex++) {
    if (sheetIndex >= sheetNames.length) break; // Пропускаем, если листа нет

    const sheetName = sheetNames[sheetIndex];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    for (let i = 1; i < data.length; i++) { // Пропускаем заголовки
      const row = data[i];
      const valueID = row[0]; // Первый столбец (предполагаемый ИД)
      const valueM = parseFloat(row[12]); // Столбец M (индекс 12)
      const valueN = parseFloat(row[13]); // Столбец N (индекс 13)

      // Проверяем, что оба значения являются числами
      if (isNaN(valueM) || isNaN(valueN)) continue;

      // Проверяем разницу
      if (valueN - valueM <= 12) {
        const difference = valueN - valueM;
        result.push(
          `Лист "${sheetName}", строка ${i + 1}: ИД="${valueID}", разница ${difference.toFixed(
            2
          )} (M=${valueM}, N=${valueN})`
        );
      }
    }
  }

  // Обработка результатов
  const resultMessage =
    result.length === 0
      ? 'Все значения столбца N - значения столбца M больше 12 на первых двух листах.'
      : result.join('\n');

  const filePath = path.join(__dirname, '..', 'Result.txt');
  fs.writeFileSync(filePath, resultMessage);

  await ctx.replyWithDocument({ source: filePath });
  fs.unlinkSync(filePath);
}

module.exports = { checkColumnDifference };