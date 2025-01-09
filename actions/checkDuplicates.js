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
    if (!row) continue;

    const key = `${row[0]}|${row[3]}|${row[4]}`; // Создаем уникальный ключ из столбцов A, D, E

    if (!duplicates.has(key)) {
      duplicates.set(key, { valuesH: new Map(), rows: [i + 1] });
      duplicates.get(key).valuesH.set(row[6], [i + 1]); // Добавляем значение H и строку
    } else {
      const existingEntry = duplicates.get(key);

      if (existingEntry.valuesH.has(row[6])) {
        const conflictingRows = existingEntry.valuesH.get(row[6]);
        conflictingRows.push(i + 1);

        result.push(
          `Лист "${sheetName}", строки ${conflictingRows.join(
            ', '
          )} имеют одинаковые значения в столбцах A, D, E, а также одинаковое значение в столбце H: "${row[6]}"`
        );
      } else {
        existingEntry.valuesH.set(row[6], [i + 1]); // Добавляем новое значение в H
        existingEntry.rows.push(i + 1); // Добавляем текущую строку в список
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