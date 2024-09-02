const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

async function checkCyrillic(ctx, workbook) {
  const sheetNames = workbook.SheetNames;
  let result = [];

  // Функция для проверки наличия кириллических символов
  function containsCyrillic(text) {
    return /[а-яА-ЯЁё]/.test(text);
  }

  // Проверяем лист "Диагностика"
  const sheetNameDiag = 'Диагностика';
  if (sheetNames.includes(sheetNameDiag)) {
    const sheetDiag = workbook.Sheets[sheetNameDiag];
    const dataDiag = xlsx.utils.sheet_to_json(sheetDiag, { header: 1 });

    for (let i = 1; i < dataDiag.length; i++) { // Начинаем с i = 1, чтобы пропустить заголовок
      const row = dataDiag[i];
      const columnsToCheckDiag = [1, 7, 19]; // Столбцы B, H, T (индексы 1, 7, 19)

      columnsToCheckDiag.forEach(colIdx => {
        const cell = row[colIdx];
        if (cell && containsCyrillic(cell.toString())) {
          const columnName = String.fromCharCode(65 + colIdx); // Преобразуем индекс в букву колонки
          result.push(`Лист ${sheetNameDiag}, строка ${i + 1}, столбец ${columnName}: ${cell}`);
        }
      });
    }
  }

  // Проверяем лист "Лечение"
  const sheetNameTreatment = 'Лечение';
  if (sheetNames.includes(sheetNameTreatment)) {
    const sheetTreatment = workbook.Sheets[sheetNameTreatment];
    const dataTreatment = xlsx.utils.sheet_to_json(sheetTreatment, { header: 1 });

    for (let i = 1; i < dataTreatment.length; i++) {
      const row = dataTreatment[i];
      const cellAF = row[31]; // Столбец AF (индекс 31)

      if (cellAF && containsCyrillic(cellAF.toString())) {
        result.push(`Лист ${sheetNameTreatment}, строка ${i + 1}, столбец AF: ${cellAF}`);
      }
    }
  }

  // Проверяем лист "Профильный специалист"
  const sheetNameSpecialist = 'Профильный специалист';
  if (sheetNames.includes(sheetNameSpecialist)) {
    const sheetSpecialist = workbook.Sheets[sheetNameSpecialist];
    const dataSpecialist = xlsx.utils.sheet_to_json(sheetSpecialist, { header: 1 });

    for (let i = 1; i < dataSpecialist.length; i++) {
      const row = dataSpecialist[i];
      const cellC = row[2]; // Столбец C (индекс 2)

      if (cellC && containsCyrillic(cellC.toString())) {
        result.push(`Лист ${sheetNameSpecialist}, строка ${i + 1}, столбец C: ${cellC}`);
      }
    }
  }

  // Обработка результатов
  if (result.length === 0) {
    await ctx.reply('Кириллические символы не найдены.');
  } else {
    const resultMessage = result.join('\n');
    const filePath = path.join(__dirname, '..', 'Result.txt');
    fs.writeFileSync(filePath, resultMessage);

    await ctx.replyWithDocument({ source: filePath });

    fs.unlinkSync(filePath);
  }
}

module.exports = { checkCyrillic };