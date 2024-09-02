const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function checkUniqueMappingsTreatment(ctx, workbook) {
  const sheetName = 'Лечение';
  const result = [];

  // Проверяем, что лист "Лечение" существует
  if (!workbook.SheetNames.includes(sheetName)) {
    await ctx.reply(`Лист "${sheetName}" не найден.`);
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  // Объекты для хранения маппинга уникальных связок
  const mappingABtoC = {}; // Для проверки (A, B) -> C
  const mappingACtoB = {}; // Для проверки (A, C) -> B

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const valueA = row[0]; // Значение столбца A
    const valueB = row[1]; // Значение столбца B
    const valueC = row[2]; // Значение столбца C

    // Пропускаем строки с пустыми обязательными полями
    if (!valueA || (!valueB && !valueC)) continue;

    // Проверка для (A, B) -> C
    if (valueB) {
      const keyAB = `${valueA}_${valueB}`;
      if (!mappingABtoC[keyAB]) {
        mappingABtoC[keyAB] = valueC;
      } else if (mappingABtoC[keyAB] !== valueC) {
        result.push(
          `Лист ${sheetName}: Связка (${valueA}, ${valueB}) связана с несколькими значениями C: "${mappingABtoC[keyAB]}" и "${valueC}" (строка ${i + 1})`
        );
      }
    }

    // Проверка для (A, C) -> B
    if (valueC) {
      const keyAC = `${valueA}_${valueC}`;
      if (!mappingACtoB[keyAC]) {
        mappingACtoB[keyAC] = valueB;
      } else if (mappingACtoB[keyAC] !== valueB) {
        result.push(
          `Лист ${sheetName}: Связка (${valueA}, ${valueC}) связана с несколькими значениями B: "${mappingACtoB[keyAC]}" и "${valueB}" (строка ${i + 1})`
        );
      }
    }
  }

  // Формируем результат и отправляем его
  if (result.length === 0) {
    await ctx.reply('Все уникальные связки соответствуют правилам.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });

    fs.unlinkSync(filePath);
  }
}

module.exports = { checkUniqueMappingsTreatment };