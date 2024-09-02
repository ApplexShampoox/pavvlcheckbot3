const { Telegraf, Markup } = require('telegraf');
const axios = require('axios');
const xlsx = require('xlsx');
require('dotenv').config();

const bot = new Telegraf(process.env.BOT_TOKEN);
const session = require('./session');
bot.use(session.middleware());

// Сюда подключать новые функции (импорт из папки actions)
const { checkIdName } = require('./actions/checkIdName');
const { checkCyrillic } = require('./actions/checkCyrillic');
const { checkRequiredColumns } = require('./actions/checkRequiredColumns');
const { checkDuplicates } = require('./actions/checkDuplicates');
const { checkYesOrEmpty } = require('./actions/checkYesOrEmpty');
const { checkICDCodes } = require('./actions/checkICDCodes');
const { checkAEFMapping } = require('./actions/checkAEFMapping');
const { checkGroupsDiag } = require('./actions/checkGroupsDiag');
const { checkServiceCodes } = require('./actions/checkServiceCodes');
const { checkUniqueMappingsTreatment } = require('./actions/checkUniqueMappingsTreatment');


//При старте бота добавляем кнопку для вызова каждой функции в формате [Markup.button.callback('Надпись на кнопке','Название импортированной выше функции')]
bot.start((ctx) => {
  ctx.reply('Выберите действие:', Markup.inlineKeyboard([
    [Markup.button.callback('Проверка идентичности ID и названия шаблонов', 'checkIdName')],
    [Markup.button.callback('Проверка на наличие кириллицы', 'checkCyrillic')],
    [Markup.button.callback('Проверка на наличие дублей диагностических услуг', 'checkDuplicates')],
    [Markup.button.callback('Проверка на однотипность заполнения полей RQKG', 'checkYesOrEmpty')],
    [Markup.button.callback('Проверка на соответсвие кодов МКБ справочнику', 'checkICDCodes')],
    [Markup.button.callback('Проверка на уникальность Группа-Порядок (Диагностика)', 'checkAEFMapping')],
    [Markup.button.callback('Проверка на уникальность Группа-Порядок (Лечение)', 'checkUniqueMappingsTreatment')],
    [Markup.button.callback('Проверка на наличие одной основной альтернативы (Диагностика)', 'checkGroupsDiag')],
    [Markup.button.callback('Проверка на валидность 804 кода', 'checkServiceCodes')],
    [Markup.button.callback('Проверка на заполненность обязательных столбцов', 'checkRequiredColumns')]


  ]));
});

//Блок добавления действий бота после нажатия кнопки (сообщение предлагающее пользователю загрузить файл и присвоение ИД этому файлу для определения, для какой функции он был загружен)
bot.action('checkIdName', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки уникальности ID и названия.');
  ctx.session.waitingForFile = 'checkIdName';
});

bot.action('checkCyrillic', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на наличие кириллицы.');
  ctx.session.waitingForFile = 'checkCyrillic';
});

bot.action('checkRequiredColumns', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на заполненность столбцов.');
  ctx.session.waitingForFile = 'checkRequiredColumns';
});

bot.action('checkDuplicates', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на наличие дублей.');
  ctx.session.waitingForFile = 'checkDuplicates';
});

bot.action('checkYesOrEmpty', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на однотипность заполнения полей RQKG.');
  ctx.session.waitingForFile = 'checkYesOrEmpty';
});

bot.action('checkICDCodes', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на соответсвие кодов МКБ справочнику.');
  ctx.session.waitingForFile = 'checkICDCodes';
});

bot.action('checkAEFMapping', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на уникальность Группа-Порядок (Диагностика).');
  ctx.session.waitingForFile = 'checkAEFMapping';
});

bot.action('checkGroupsDiag', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки  на наличие одной основной альтернативы (Диагностика).');
  ctx.session.waitingForFile = 'checkGroupsDiag';
});

bot.action('checkUniqueMappingsTreatment', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на уникальность Группа-Порядок (Лечение).');
  ctx.session.waitingForFile = 'checkUniqueMappingsTreatment';
});


bot.action('checkServiceCodes', (ctx) => {
  ctx.reply('Загрузите xlsx файл для проверки на валидность 804 кода.');
  ctx.session.waitingForFile = 'checkServiceCodes';
});


//Добавление кнопки возврата в меню выбора функции (также при добавлении новой функции сюда нужно добавить кнопку по примеру выше)
bot.action('backToMenu', (ctx) => {
  ctx.reply('Выберите действие:', Markup.inlineKeyboard([
    [Markup.button.callback('Проверка идентичности принадлежности к одному шаблону/группе ', 'checkIdName')],
    [Markup.button.callback('Проверка на наличие кириллицы', 'checkCyrillic')],
    [Markup.button.callback('Проверка на наличие дублей диагностических услуг', 'checkDuplicates')],
    [Markup.button.callback('Проверка на однотипность заполнения полей RQKG', 'checkYesOrEmpty')],
    [Markup.button.callback('Проверка на соответсвие кодов МКБ справочнику', 'checkICDCodes')],
    [Markup.button.callback('Проверка на уникальность Группа-Порядок (Диагностика)', 'checkAEFMapping')],
    [Markup.button.callback('Проверка на уникальность Группа-Порядок (Лечение)', 'checkUniqueMappingsTreatment')],
    [Markup.button.callback('Проверка на наличие одной основной альтернативы (Диагностика)', 'checkGroupsDiag')],
    [Markup.button.callback('Проверка на валидность 804 кода', 'checkServiceCodes')],
    [Markup.button.callback('Проверка на заполненность обязательных столбцов', 'checkRequiredColumns')]



  ]));
  ctx.session.waitingForFile = false; // Сброс состояния ожидания файла при возврате в меню
});

bot.on('document', async (ctx) => {
  if (ctx.session.waitingForFile) {
    const fileId = ctx.message.document.file_id;
    const fileLink = await bot.telegram.getFileLink(fileId);

    try {
      const response = await axios.get(fileLink, { responseType: 'arraybuffer' });
      const buffer = Buffer.from(response.data, 'binary');
      const workbook = xlsx.read(buffer, { type: 'buffer' });

      //Определение системой какую функцию вызывать на основе ИД загруженного пользователем файла
      switch (ctx.session.waitingForFile) {
        case 'checkIdName':
          await checkIdName(ctx, workbook);
          break;
        case 'checkCyrillic':
          await checkCyrillic(ctx, workbook);
          break;
        case 'checkRequiredColumns':
          await checkRequiredColumns(ctx, workbook);
          break;
        case 'checkDuplicates':
          await checkDuplicates(ctx, workbook);
          break;
        case 'checkYesOrEmpty':
          await checkYesOrEmpty(ctx, workbook);
          break;
        case 'checkICDCodes':
          await checkICDCodes(ctx, workbook);
          break;
        case 'checkAEFMapping':
          await checkAEFMapping(ctx, workbook);
          break;
        case 'checkGroupsDiag':
          await checkGroupsDiag(ctx, workbook);
          break;
        case 'checkUniqueMappingsTreatment':
          await checkUniqueMappingsTreatment(ctx, workbook);
          break;
        case 'checkServiceCodes':
          await checkServiceCodes(ctx, workbook);
          break;



        // Добавляйте новые функции здесь
        default:
          ctx.reply('Неизвестное действие. Пожалуйста, попробуйте снова.');
          break;
      }
      // Сообщение после обработки файла
      ctx.reply('Загрузите следующий xlsx файл или вернитесь в меню.', Markup.inlineKeyboard([
        Markup.button.callback('Вернуться в меню', 'backToMenu')
      ]));

      // Установка состояния ожидания файла для текущей функции
      ctx.session.waitingForFile = {
        'checkIdName': 'checkIdName',
        'checkCyrillic': 'checkCyrillic',
        'checkDuplicates': 'checkDuplicates',
        'checkYesOrEmpty': 'checkYesOrEmpty',
        'checkICDCodes': 'checkICDCodes',
        'checkAEFMapping': 'checkAEFMapping',
        'checkGroupsDiag': 'checkGroupsDiag',
        'checkUniqueMappingsTreatment': 'checkUniqueMappingsTreatment',
        'checkServiceCodes': 'checkServiceCodes',
        'checkRequiredColumns': 'checkRequiredColumns'


      }[ctx.session.waitingForFile];

    } catch (error) {
      console.error('Ошибка при обработке файла:', error);
      ctx.reply('Произошла ошибка при обработке файла. Пожалуйста, попробуйте снова.');
    }
  } else {
    ctx.reply('Пожалуйста, выберите действие сначала.');
  }
});

bot.launch();