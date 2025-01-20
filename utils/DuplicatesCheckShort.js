//Проверяет на наличие дублей 1 лист 1-2 столбцы по отдельности и вместе
const XLSX = require('xlsx');

// Функция для поиска дубликатов в массиве
const findDuplicates = (array) => {
  const counts = {};
  const duplicates = [];

  array.forEach((item) => {
    const key = JSON.stringify(item); // Учитываем возможность сложных ключей
    counts[key] = (counts[key] || 0) + 1;
    if (counts[key] === 2) duplicates.push(item);
  });

  return duplicates;
};

// Основная функция проверки дубликатов
const checkDuplicates = (filePath) => {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Берем первый лист
    const sheet = workbook.Sheets[sheetName];

    // Конвертируем лист в массив объектов
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Массив массивов
    if (data.length === 0) {
      console.log('Таблица пуста.');
      return;
    }

    const column1 = data.map(row => row[0]).slice(1); // Первый столбец (без заголовков)
    const column2 = data.map(row => row[1]).slice(1); // Второй столбец (без заголовков)
    const combined = data.slice(1).map(row => `${row[0]}|${row[1]}`); // Комбинация столбцов 1 и 2

    // Проверка на пустые строки
    const emptyRows = [];
    data.slice(1).forEach((row, index) => {
      if (!row[0] || !row[1]) emptyRows.push(index + 2); // Добавляем 2 для корректного номера строки
    });

    const duplicatesCol1 = findDuplicates(column1);
    const duplicatesCol2 = findDuplicates(column2);
    const duplicatesCombined = findDuplicates(combined);

    console.log('Дубликаты в первом столбце:', duplicatesCol1);
    console.log('Дубликаты во втором столбце:', duplicatesCol2);
    console.log('Дубликаты в комбинации первого и второго столбцов:', duplicatesCombined);

    if (emptyRows.length > 0) {
      console.log('Пустые строки (оба столбца должны быть заполнены):', emptyRows);
    } else {
      console.log('Все строки заполнены.');
    }
  } catch (error) {
    console.error('Ошибка при обработке файла:', error.message);
  }
};

// Путь к файлу
const filePath = '../inptut.xlsx'; // Укажите путь к вашему файлу
checkDuplicates(filePath);