//Забирает первый столбец первого листа и превращает значения в массив строк

const fs = require('fs');
const XLSX = require('xlsx');

// Функция для обработки файла
function extractFirstColumnToJS(xlsxFilePath, jsFilePath) {
  try {
    // Читаем файл Excel
    const workbook = XLSX.readFile(xlsxFilePath);

    // Получаем первый лист
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Преобразуем лист в JSON
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Извлекаем первый столбец (если есть данные)
    const firstColumn = sheetData
      .map(row => row[0])
      .filter(value => value !== undefined)
      .map(value => String(value)); // Преобразуем всё в строки

    // Генерируем содержимое JS файла
    const jsContent = `const data = ${JSON.stringify(firstColumn)};\n\nmodule.exports = data;\n`;

    // Сохраняем в файл
    fs.writeFileSync(jsFilePath, jsContent, 'utf8');

    console.log(`Данные успешно сохранены в файл: ${jsFilePath}`);
  } catch (error) {
    console.error('Произошла ошибка:', error.message);
  }
}

// Укажите пути к файлам
const xlsxFilePath = '../input.xlsx'; // Исходный Excel файл
const jsFilePath = '../output.js';   // Итоговый JS файл

// Запуск функции
extractFirstColumnToJS(xlsxFilePath, jsFilePath);