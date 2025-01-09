const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function checkGroupsDiag(ctx, workbook) {
  const sheetName = 'Диагностика';
  const result = [];

  // Проверяем, что лист "Диагностика" существует
  if (!workbook.SheetNames.includes(sheetName)) {
    await ctx.reply(`Лист "${sheetName}" не найден.`);
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  // Объект для хранения групп по значениям A, D, E, I и J
  const groups = {};

  for (let i = 1; i < data.length; i++) { // Пропускаем первую строку с заголовками
    const row = data[i];
    const valueA = String(row[0] || '').trim(); // Значение в столбце A
    const valueD = String(row[3] || '').trim(); // Значение в столбце D
    const valueE = String(row[4] || '').trim(); // Значение в столбце E
    const valueI = String(row[8] || '').trim(); // Значение в столбце I
    const valueJ = String(row[9] || '').trim(); // Значение в столбце J

    const groupKey = `${valueA}_${valueD}_${valueE}_${valueI}`;

    // Пропускаем строки, где I пустой
    if (!valueI) continue;

    // Инициализация группы, если не существует
    if (!groups[groupKey]) {
      groups[groupKey] = {
        rows: [],
        filledJCount: 0,
      };
    }

    // Добавляем строку в группу
    groups[groupKey].rows.push({ rowNumber: i + 1, valueJ });

    // Увеличиваем счётчик строк с заполненным столбцом J
    if (valueJ) {
      groups[groupKey].filledJCount++;
    }
  }

  // Проверка групп на наличие только одной строки с заполненным столбцом J
  Object.keys(groups).forEach((groupKey) => {
    const group = groups[groupKey];

    // Проводим проверку только для групп, где столбец I не пустой
    if (group.filledJCount !== 1) {
      const rows = group.rows.map(r => `строка ${r.rowNumber}`).join(', ');
      result.push(`Несоответствие на листе ${sheetName}: группа с значениями A, D, E, I "${groupKey.replace(/_/g, ', ')}" должна содержать одну строку с заполненным J, но найдено ${group.filledJCount} (строки: ${rows})`);
    }
  });

  // Обработка результатов
  if (result.length === 0) {
    await ctx.reply('Все группы удовлетворяют условию.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });
    fs.unlinkSync(filePath);
  }
}

module.exports = { checkGroupsDiag };