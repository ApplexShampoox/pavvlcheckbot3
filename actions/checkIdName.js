const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

async function checkIdName(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];

  function checkSheet(sheetName) {
    if (!sheetNames.includes(sheetName)) {
      result.push(`Лист ${sheetName} отсутствует.`);
      return; // Пропустить, если листа нет
    }

    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Объекты для проверки уникальности
    let valueAToC = {};
    let valueCToA = {};

    for (let i = 1; i < data.length; i++) { // Пропускаем первую строку с заголовками
      const row = data[i];
      const valueA = row[0]; // Столбец A
      const valueC = row[2]; // Столбец C

      // Пропуск пустых строк или строк с пустыми значениями
      if (!valueA || !valueC) continue;

      // Проверка уникальности по valueAToC
      if (valueAToC[valueA]) {
        if (valueAToC[valueA] !== valueC) {
          result.push(
            `Несоответствие на листе ${sheetName}: значение "${valueA}" в столбце A связано с несколькими значениями столбца C: "${valueAToC[valueA]}" и "${valueC}" (строка ${i + 1})`
          );
        }
      } else {
        valueAToC[valueA] = valueC;
      }

      // Проверка уникальности по valueCToA
      if (valueCToA[valueC]) {
        if (valueCToA[valueC] !== valueA) {
          result.push(
            `Несоответствие на листе ${sheetName}: значение "${valueC}" в столбце C связано с несколькими значениями столбца A: "${valueCToA[valueC]}" и "${valueA}" (строка ${i + 1})`
          );
        }
      } else {
        valueCToA[valueC] = valueA;
      }
    }
  }

  // Проверяем лист "Диагностика"
  checkSheet('Диагностика');

  // Обработка результатов
  const resultMessage =
    result.length === 0
      ? 'Для каждого уникального значения в столбце A существует только одно уникальное значение в столбце C, и наоборот.'
      : result.join('\n');

  const filePath = path.join(__dirname, '..', 'Result.txt');
  fs.writeFileSync(filePath, resultMessage);

  await ctx.replyWithDocument({ source: filePath });
  fs.unlinkSync(filePath);
}

module.exports = { checkIdName };