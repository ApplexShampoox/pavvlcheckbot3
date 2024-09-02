const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function checkAEFMapping(ctx, workbook) {
  const sheetName = 'Диагностика';
  const result = [];

  // Проверяем, что лист "Диагностика" существует
  if (!workbook.SheetNames.includes(sheetName)) {
    await ctx.reply(`Лист "${sheetName}" не найден.`);
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  // Объекты для хранения маппинга уникальных связок
  const mappingAEFtoD = {}; // Для проверки (A, D, E) -> F
  const mappingADFtoE = {}; // Для проверки (A, D, F) -> E

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const valueA = row[0];
    const valueD = row[3];
    const valueE = row[4];
    const valueF = row[5];

    // Пропускаем строки с пустыми обязательными полями
    if (!valueA || !valueD) continue;

    // Проверка для (A, D, E) -> F
    if (valueE) {
      const keyAEF = `${valueA}_${valueD}_${valueE}`;
      if (!mappingAEFtoD[keyAEF]) {
        mappingAEFtoD[keyAEF] = valueF;
      } else if (mappingAEFtoD[keyAEF] !== valueF) {
        result.push(
          `Лист ${sheetName}: Связка (${valueA}, ${valueD}, ${valueE}) связана с несколькими значениями F: "${mappingAEFtoD[keyAEF]}" и "${valueF}" (строка ${i + 1})`
        );
      }
    }

    // Проверка для (A, D, F) -> E
    if (valueF) {
      const keyADF = `${valueA}_${valueD}_${valueF}`;
      if (!mappingADFtoE[keyADF]) {
        mappingADFtoE[keyADF] = valueE;
      } else if (mappingADFtoE[keyADF] !== valueE) {
        result.push(
          `Лист ${sheetName}: Связка (${valueA}, ${valueD}, ${valueF}) связана с несколькими значениями E: "${mappingADFtoE[keyADF]}" и "${valueE}" (строка ${i + 1})`
        );
      }
    }
  }

  // Формируем результат и отправляем его
  if (result.length === 0) {
    await ctx.reply('Все уникальные связки уникальны.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });

    fs.unlinkSync(filePath);
  }
}

module.exports = { checkAEFMapping };