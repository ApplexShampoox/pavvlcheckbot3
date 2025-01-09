const xlsx = require('xlsx');

function checkCyrillicInFirstColumn(filePath) {
  try {
    // Читаем Excel-файл
    const workbook = xlsx.readFile(filePath);

    // Получаем первый лист
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Преобразуем лист в массив строк
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    console.log(`Проверяем первый столбец на листе "${sheetName}"...`);

    // Проверка на кириллические символы
    const cyrillicRegex = /[а-яА-ЯёЁ]/; // Регулярное выражение для поиска кириллицы
    let found = false;

    for (let i = 0; i < data.length; i++) {
      const cellValue = data[i][0]; // Первый столбец (индекс 0)
      if (cellValue && cyrillicRegex.test(cellValue)) {
        console.log(`Строка ${i + 1}: "${cellValue}" содержит кириллические символы.`);
        found = true;
      }
    }

    if (!found) {
      console.log('Кириллические символы не найдены в первом столбце.');
    }
  } catch (error) {
    console.error('Ошибка при обработке файла:', error.message);
  }
}

// Укажите путь к вашему Excel-файлу
const filePath = './Услуги По 804Н.xlsx';
checkCyrillicInFirstColumn(filePath);