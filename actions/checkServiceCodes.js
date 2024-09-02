const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const allowedValues = require('./../804Services.js'); // Подключаем массив допустимых значений из values.js

async function checkServiceCodes(ctx, workbook) {
  const result = [];

  // Проверяем лист "Диагностика"
  const diagSheetName = 'Диагностика';
  if (workbook.SheetNames.includes(diagSheetName)) {
    const diagSheet = workbook.Sheets[diagSheetName];
    const diagData = xlsx.utils.sheet_to_json(diagSheet, { header: 1 });

    for (let i = 1; i < diagData.length; i++) {
      const serviceCode = diagData[i][7]; // Столбец H
      if (serviceCode !== undefined && !allowedValues.includes(String(serviceCode).trim())) {
        result.push(
          `Лист ${diagSheetName}: значение "${serviceCode}" в столбце H (строка ${i + 1}) не входит в список допустимых значений.`
        );
      }
    }
  } else {
    await ctx.reply(`Лист "${diagSheetName}" не найден.`);
  }

  // Проверяем лист "Профильный специалист"
  const treatmentSheetName = 'Профильный специалист';
  if (workbook.SheetNames.includes(treatmentSheetName)) {
    const treatmentSheet = workbook.Sheets[treatmentSheetName];
    const treatmentData = xlsx.utils.sheet_to_json(treatmentSheet, { header: 1 });

    for (let i = 1; i < treatmentData.length; i++) {
      const serviceCode = treatmentData[i][2]; // Столбец C
      if (serviceCode !== undefined && !allowedValues.includes(String(serviceCode).trim())) {
        result.push(
          `Лист ${treatmentSheetName}: значение "${serviceCode}" в столбце C (строка ${i + 1}) не входит в список допустимых значений.`
        );
      }
    }
  } else {
    await ctx.reply(`Лист "${treatmentSheetName}" не найден.`);
  }

  // Формируем результат и отправляем его
  if (result.length === 0) {
    await ctx.reply('Все значения соответствуют списку допустимых.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });

    fs.unlinkSync(filePath);
  }
}

module.exports = { checkServiceCodes };