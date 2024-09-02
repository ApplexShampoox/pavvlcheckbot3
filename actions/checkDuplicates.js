const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

async function checkDuplicates(ctx, workbook) {
  const sheetName = 'Диагностика'; // Лист для проверки
  let result = [];

  if (!workbook.SheetNames.includes(sheetName)) {
    await ctx.reply(`Лист "${sheetName}" не найден в файле.`);
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  const duplicates = new Map();

  // Пропускаем первую строку с заголовками и начинаем с первой строки данных
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = `${row[0]}|${row[3]}|${row[13]}`; // Создаем уникальный ключ из столбцов A, D, N

    if (!duplicates.has(key)) {
      duplicates.set(key, { valueH: row[7], rows: [i + 1] });
    } else {
      const existingEntry = duplicates.get(key);
      existingEntry.rows.push(i + 1);

      // Если значение в столбце H такое же, как и в предыдущей строке с тем же ключом
      if (existingEntry.valueH === row[7]) {
        result.push(
          `Лист "${sheetName}", строки ${existingEntry.rows.join(
            ', '
          )} имеют одинаковые значения в столбцах A, D, N, но также одинаковое значение в столбце H: "${row[7]}"`
        );
      }
    }
  }

  // Обработка результатов
  if (result.length === 0) {
    await ctx.reply('Дубликаты не найдены.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });

    fs.unlinkSync(filePath);
  }
}

module.exports = { checkDuplicates };