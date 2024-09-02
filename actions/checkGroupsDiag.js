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

  // Объект для хранения групп по значениям A, D и J
  const groups = {};

  for (let i = 1; i < data.length; i++) { // Пропускаем первую строку с заголовками
    const row = data[i];
    const valueA = row[0]; // Значение в столбце A
    const valueD = row[3]; // Значение в столбце D
    const valueJ = row[9]; // Значение в столбце J
    const valueK = row[10]; // Значение в столбце K

    // Пропускаем строки, где A, D или J не заполнены
    if (!valueA || !valueD || !valueJ) continue;

    const groupKey = `${valueA}_${valueD}_${valueJ}`;

    // Инициализация группы, если не существует
    if (!groups[groupKey]) {
      groups[groupKey] = {
        rows: [],
        filledKCount: 0,
      };
    }

    // Добавляем строку в группу
    groups[groupKey].rows.push({ rowNumber: i + 1, valueK });

    // Увеличиваем счетчик заполненных K
    if (valueK && valueK.trim() !== '') {
      groups[groupKey].filledKCount++;
    }
  }

  // Проверка групп на наличие только одной строки с заполненным столбцом K
  Object.keys(groups).forEach((groupKey) => {
    const group = groups[groupKey];

    if (group.filledKCount !== 1) {
      const rows = group.rows.map(r => `строка ${r.rowNumber}`).join(', ');
      result.push(`Несоответствие на листе ${sheetName}: группа с значениями A, D и J "${groupKey.replace(/_/g, ', ')}" должна содержать одну строку с заполненным K, но найдено ${group.filledKCount} (строки: ${rows})`);
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