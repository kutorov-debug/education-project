import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../outputs/futures-learning-lab",
);
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const overview = workbook.worksheets.add("Обзор");
const contract = workbook.worksheets.add("Контракт");
const calculator = workbook.worksheets.add("Калькулятор");
const journal = workbook.worksheets.add("Журнал сделок");
const progress = workbook.worksheets.add("Прогресс");
const sources = workbook.worksheets.add("Источники");
const checks = workbook.worksheets.add("Проверки");

const COLORS = {
  navy: "#12355B",
  blue: "#0B5CAB",
  lightBlue: "#E8F1FB",
  green: "#2E7D32",
  lightGreen: "#E8F5E9",
  amber: "#A86400",
  lightAmber: "#FFF4D6",
  red: "#B42318",
  lightRed: "#FDECEC",
  gray: "#5B6573",
  lightGray: "#F3F5F7",
  border: "#CBD3DC",
  white: "#FFFFFF",
  black: "#111827",
};

const title = (sheet, range, text) => {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 18 },
    verticalAlignment: "center",
  };
};

const section = (sheet, range, text) => {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white, size: 11 },
    verticalAlignment: "center",
  };
};

const header = (range) => {
  range.format = {
    fill: COLORS.lightBlue,
    font: { bold: true, color: COLORS.navy },
    borders: { preset: "outside", style: "thin", color: COLORS.border },
    verticalAlignment: "center",
    wrapText: true,
  };
};

const note = (range) => {
  range.format = {
    fill: COLORS.lightAmber,
    font: { color: COLORS.amber },
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: "#E7C66B" },
    verticalAlignment: "center",
  };
};

const input = (range) => {
  range.format = {
    fill: "#F7FBFF",
    font: { color: "#0000FF" },
    borders: { preset: "outside", style: "thin", color: COLORS.border },
  };
};

const formula = (range) => {
  range.format = {
    fill: COLORS.white,
    font: { color: COLORS.black },
    borders: { preset: "outside", style: "thin", color: COLORS.border },
  };
};

for (const sheet of [overview, contract, calculator, journal, progress, sources, checks]) {
  sheet.showGridLines = false;
}

// Обзор
title(overview, "A1:H2", "Учебная лаборатория: фьючерсы");
overview.getRange("A3:H3").merge();
overview.getRange("A3:H3").values = [[
  "Бумажная практика и расчёты. Это не брокерский терминал, не торговый сигнал и не доказательство будущей прибыли.",
]];
note(overview.getRange("A3:H3"));
overview.getRange("A5:H5").values = [[
  "Бумажных сделок", null, "Средний результат после издержек", null,
  "Доля без превышения риск-бюджета", null, "Общий статус модели", null,
]];
overview.getRange("A5:H5").format = {
  fill: COLORS.lightGray,
  font: { bold: true, color: COLORS.gray },
  wrapText: true,
  verticalAlignment: "center",
};
overview.getRange("A6:B7").merge();
overview.getRange("C6:D7").merge();
overview.getRange("E6:F7").merge();
overview.getRange("G6:H7").merge();
overview.getRange("A6").formulas = [["=COUNTIF('Журнал сделок'!$B$7:$B$106,\"<>\")"]];
overview.getRange("C6").formulas = [["=IFERROR(AVERAGE('Журнал сделок'!$T$7:$T$106),0)"]];
overview.getRange("E6").formulas = [["=IF(A6=0,0,COUNTIF('Журнал сделок'!$U$7:$U$106,\"Нет\")/A6)"]];
overview.getRange("G6").formulas = [["='Проверки'!$B$12"]];
overview.getRange("A6:H7").format = {
  fill: COLORS.white,
  font: { bold: true, color: COLORS.navy, size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.border },
};
overview.getRange("G6").format.font = { bold: true, color: COLORS.navy, size: 13 };
overview.getRange("C6").format.numberFormat = "#,##0;[Red](#,##0);-";
overview.getRange("E6").format.numberFormat = "0%";

section(overview, "A9:H9", "Как пользоваться");
overview.getRange("A10:H13").values = [
  ["1", "Контракт", "Перенесите только проверенные поля из карточки MOEX ISS вместе с URL и временем среза.", null, null, null, null, null],
  ["2", "Калькулятор", "Отделяйте риск-бюджет, планируемый убыток до стопа и фактический результат.", null, null, null, null, null],
  ["3", "Журнал", "Записывайте план до сделки и учебное исполнение после сделки в разные столбцы.", null, null, null, null, null],
  ["4", "Прогресс", "Повышайте статус только по самостоятельному действию и отсроченной проверке.", null, null, null, null, null],
];
overview.getRange("B10:B13").format.font = { bold: true, color: COLORS.blue };
overview.getRange("C10:H13").merge(true);
overview.getRange("A10:H13").format.wrapText = true;
overview.getRange("A10:H13").format.borders = { preset: "inside", style: "thin", color: COLORS.border };

section(overview, "A15:H15", "Готовность к следующему этапу");
overview.getRange("A16:H16").values = [["Ворота", null, null, "Статус", "Свидетельство", null, null, "Следующая проверка"]];
overview.getRange("A16:C16").merge();
overview.getRange("E16:G16").merge();
header(overview.getRange("A16:H16"));
const gates = [
  ["Паспорт контракта и единицы", "='Прогресс'!$B$8", "='Прогресс'!$C$8", "=IF('Прогресс'!$E$8=\"\",\"\",'Прогресс'!$E$8)"],
  ["P&L двумя способами", "='Прогресс'!$B$9", "='Прогресс'!$C$9", "=IF('Прогресс'!$E$9=\"\",\"\",'Прогресс'!$E$9)"],
  ["ГО, номинал, капитал и риск-бюджет", "='Прогресс'!$B$10", "='Прогресс'!$C$10", "=IF('Прогресс'!$E$10=\"\",\"\",'Прогресс'!$E$10)"],
  ["Стоп, исполнение и издержки", "='Прогресс'!$B$11", "='Прогресс'!$C$11", "=IF('Прогресс'!$E$11=\"\",\"\",'Прогресс'!$E$11)"],
  ["Серия бумажных решений", "='Прогресс'!$B$12", "='Прогресс'!$C$12", "=IF('Прогресс'!$E$12=\"\",\"\",'Прогресс'!$E$12)"],
  ["Статистика и ограничения выборки", "='Прогресс'!$B$13", "='Прогресс'!$C$13", "=IF('Прогресс'!$E$13=\"\",\"\",'Прогресс'!$E$13)"],
];
for (let i = 0; i < gates.length; i += 1) {
  const row = 17 + i;
  overview.getRange(`A${row}:C${row}`).merge();
  overview.getRange(`E${row}:G${row}`).merge();
  overview.getRange(`A${row}`).values = [[gates[i][0]]];
  overview.getRange(`D${row}`).formulas = [[gates[i][1]]];
  overview.getRange(`E${row}`).formulas = [[gates[i][2]]];
  overview.getRange(`H${row}`).formulas = [[gates[i][3]]];
}
overview.getRange("A17:H22").format = {
  borders: { preset: "inside", style: "thin", color: COLORS.border },
  wrapText: true,
  verticalAlignment: "center",
};
overview.getRange("H17:H22").format.numberFormat = "yyyy-mm-dd";
overview.getRange("D17:D22").conditionalFormats.add("containsText", {
  text: "перенос подтверждён",
  format: { fill: COLORS.lightGreen, font: { color: COLORS.green, bold: true } },
});
overview.getRange("A23:H23").merge();
overview.getRange("A23:H23").values = [["Диаграмма появится после первой заполненной бумажной сделки."]];
overview.getRange("A23:H23").format = {
  fill: COLORS.lightGray,
  font: { italic: true, color: COLORS.gray },
  horizontalAlignment: "center",
};

const resultChart = overview.charts.add("bar", {
  chartType: "bar",
  title: "Результат бумажных сделок после издержек, ₽",
  hasLegend: false,
});
const resultSeries = resultChart.series.add("Результат после издержек");
resultSeries.categoryFormula = "'Журнал сделок'!$A$7:$A$26";
resultSeries.formula = "'Журнал сделок'!$T$7:$T$26";
resultSeries.fill = COLORS.blue;
resultChart.title = "Результат бумажных сделок после издержек, ₽";
resultChart.hasLegend = false;
resultChart.yAxis = { numberFormatCode: "#,##0;[Red](#,##0)" };
resultChart.setPosition("A25", "H39");

// Контракт
title(contract, "A1:D2", "Паспорт фьючерсного контракта");
contract.getRange("A3:D3").merge();
contract.getRange("A3:D3").values = [[
  "Синие значения переносите из текущего ответа MOEX ISS. Всегда сохраняйте SECID, URL и время среза.",
]];
note(contract.getRange("A3:D3"));
contract.getRange("A5:D5").values = [["Поле", "Значение", "Единица", "Пояснение / источник"]];
header(contract.getRange("A5:D5"));
const contractRows = [
  ["SECID", null, "код", "Точный идентификатор активной серии"],
  ["Время среза", null, "дата и время", "Поле SYSTIME и время получения ответа"],
  ["Source URL", null, "URL", "Прямой URL ответа MOEX ISS"],
  ["Краткое название", null, "текст", "SHORTNAME"],
  ["Базовый актив", null, "код", "ASSETCODE"],
  ["Наблюдаемая цена LAST", null, "пункты котировки", "Не гарантированная цена исполнения"],
  ["Предыдущая расчётная цена", null, "пункты котировки", "PREVSETTLEPRICE"],
  ["Минимальный шаг цены", null, "пункты котировки", "MINSTEP"],
  ["Стоимость шага", null, "₽/контракт", "STEPPRICE; проверить спецификацию"],
  ["Объём лота", null, "единиц базового актива", "LOTVOLUME; смысл зависит от контракта"],
  ["Гарантийное обеспечение", null, "₽/контракт", "INITIALMARGIN; не цена и не максимум убытка"],
  ["Последний день торговли", null, "дата", "LASTTRADEDATE"],
  ["Дата исполнения", null, "дата", "LASTDELDATE"],
  ["Номинал 1 контракта", null, "₽", "Вводить только после проверки формулы спецификации"],
];
contract.getRange("A6:D19").values = contractRows;
input(contract.getRange("B6:B19"));
contract.getRange("B6:B19").format.font = { color: "#0000FF" };
contract.getRange("A6:D19").format.wrapText = true;
contract.getRange("A6:D19").format.borders = { preset: "inside", style: "thin", color: COLORS.border };
contract.getRange("B11:B16").format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
contract.getRange("B17:B18").format.numberFormat = "yyyy-mm-dd";
contract.getRange("B19").format.numberFormat = "#,##0;[Red](#,##0);-";
section(contract, "A21:D21", "Терминологические проверки");
contract.getRange("A22:D25").values = [
  ["ГО", "Текущий параметр обеспечения", null, "Не равно цене контракта"],
  ["LAST", "Наблюдаемое значение", null, "Не гарантирует исполнение"],
  ["Риск-бюджет", "Плановый лимит до сделки", null, "Не максимально возможный убыток"],
  ["Стоп", "Условие на отправку/исполнение заявки", null, "Не гарантирует цену выхода"],
];
contract.getRange("B22:C25").merge(true);
contract.getRange("A22:D25").format.wrapText = true;
contract.getRange("A22:D25").format.borders = { preset: "inside", style: "thin", color: COLORS.border };

// Калькулятор
title(calculator, "A1:D2", "Калькулятор риска и результата");
calculator.getRange("A3:D3").merge();
calculator.getRange("A3:D3").values = [[
  "Не называйте риск-бюджет максимальным убытком: стоп и исполнение не гарантируют этот предел.",
]];
note(calculator.getRange("A3:D3"));
section(calculator, "A5:D5", "План до сделки");
calculator.getRange("A6:D15").values = [
  ["Капитал", null, "₽", "Редактируемый вход"],
  ["Риск-бюджет", null, "₽", "Плановый лимит до сделки"],
  ["Резерв к ГО", 1.2, "x", "Учебное допущение; согласовать отдельно"],
  ["Плановая цена входа", null, "пункты", "До сделки"],
  ["Плановый стоп", null, "пункты", "Не гарантированная цена исполнения"],
  ["Направление", null, "Лонг/Шорт", "Знак позиции"],
  ["Комиссии на круг / контракт", null, "₽", "Биржевые + брокерские"],
  ["Плановое проскальзывание / контракт", null, "₽", "Явное учебное допущение"],
  ["Число контрактов — вручную", null, "контрактов", "Выбор ученика"],
  ["Число контрактов — расчётное", null, "контрактов", "Минимум ограничений по риску и ГО"],
];
input(calculator.getRange("B6:B14"));
calculator.getRange("B15").formulas = [["=IF(OR(B22=\"\",B23=\"\"),\"\",MIN(B22,B23))"]];
formula(calculator.getRange("B15"));
calculator.getRange("B6:B7").format.numberFormat = "#,##0;[Red](#,##0);-";
calculator.getRange("B8").format.numberFormat = "0.0x";
calculator.getRange("B9:B10").format.numberFormat = "#,##0.00";
calculator.getRange("B12:B13").format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
calculator.getRange("B14:B15").format.numberFormat = "0";
calculator.getRange("B11").dataValidation = { rule: { type: "list", values: ["Лонг", "Шорт"] } };

section(calculator, "A17:D17", "Расчёт плана");
calculator.getRange("A18:D23").values = [
  ["Дистанция до стопа", null, "пунктов", "ABS(вход − стоп)"],
  ["Убыток до стопа / контракт", null, "₽", "До комиссий и проскальзывания"],
  ["Планируемый убыток до стопа", null, "₽", "С учётом плановых издержек"],
  ["Запас риск-бюджета", null, "₽", "Риск-бюджет − планируемый убыток"],
  ["Лимит контрактов по риск-бюджету", null, "контрактов", "Не использовать до введения метода"],
  ["Лимит контрактов по ГО", null, "контрактов", "Текущее ГО × резерв"],
];
calculator.getRange("B18:B23").formulas = [
  ["=IF(OR(B9=\"\",B10=\"\"),\"\",ABS(B9-B10))"],
  ["=IF(OR(B18=\"\",'Контракт'!$B$13<=0,'Контракт'!$B$14<=0),\"\",B18/'Контракт'!$B$13*'Контракт'!$B$14)"],
  ["=IF(OR(B19=\"\",B14=\"\"),\"\",B19*B14+(B12+B13)*B14)"],
  ["=IF(OR(B7=\"\",B20=\"\"),\"\",B7-B20)"],
  ["=IF(OR(B7=\"\",B19=\"\",B12=\"\",B13=\"\"),\"\",IFERROR(INT(B7/(B19+B12+B13)),\"\"))"],
  ["=IF(OR(B6=\"\",'Контракт'!$B$16<=0,B8<=0),\"\",IFERROR(INT(B6/('Контракт'!$B$16*B8)),\"\"))"],
];
formula(calculator.getRange("B18:B23"));
calculator.getRange("B18").format.numberFormat = "#,##0.00";
calculator.getRange("B19:B21").format.numberFormat = "#,##0;[Red](#,##0);-";
calculator.getRange("B22:B23").format.numberFormat = "0";

section(calculator, "A25:D25", "Учебное исполнение после сделки");
calculator.getRange("A26:D35").values = [
  ["Фактическая цена входа", null, "пункты", "После учебного исполнения"],
  ["Фактическая цена выхода", null, "пункты", "После учебного исполнения"],
  ["Фактические комиссии", null, "₽", "Не прятать внутри цены"],
  ["Фактическое проскальзывание", null, "₽", "Не прятать внутри цены"],
  ["Валовый результат", null, "₽", "До издержек"],
  ["Фактический результат", null, "₽", "После комиссий и проскальзывания"],
  ["Фактический убыток", null, "₽", "Только отрицательная часть результата"],
  ["Превышен риск-бюджет?", null, "Да/Нет", "Сравнение, не обещание предела"],
  ["Максимально возможный убыток", null, "—", "Не определяется автоматически"],
  ["Контроль единиц", null, "OK/Проверить", "MINSTEP и STEPPRICE заполнены"],
];
input(calculator.getRange("B26:B29"));
calculator.getRange("B30:B35").formulas = [
  ["=IF(OR(B26=\"\",B27=\"\",B14=\"\",B11=\"\",'Контракт'!$B$13<=0,'Контракт'!$B$14<=0),\"\",IF(B11=\"Лонг\",1,-1)*(B27-B26)/'Контракт'!$B$13*'Контракт'!$B$14*B14)"],
  ["=IF(OR(B30=\"\",B28=\"\",B29=\"\"),\"\",B30-B28-B29)"],
  ["=IF(B31=\"\",\"\",MAX(0,-B31))"],
  ["=IF(OR(B32=\"\",B7=\"\"),\"\",IF(B32>B7,\"Да\",\"Нет\"))"],
  ["=\"Не определён: нужен отдельный сценарий и границы инструмента\""],
  ["=IF(AND('Контракт'!$B$13>0,'Контракт'!$B$14>0),\"OK\",\"Проверить\")"],
];
formula(calculator.getRange("B30:B35"));
calculator.getRange("B26:B32").format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
calculator.getRange("B33").conditionalFormats.add("containsText", {
  text: "Да",
  format: { fill: COLORS.lightRed, font: { color: COLORS.red, bold: true } },
});

// Журнал сделок
title(journal, "A1:W2", "Журнал бумажных сделок");
journal.getRange("A3:W3").merge();
journal.getRange("A3:W3").values = [[
  "Заполняйте план до сделки и фактическую часть после исполнения модели. Синие поля — ввод; формулы — чёрные; ссылки на лист «Контракт» — зелёные.",
]];
note(journal.getRange("A3:W3"));
journal.getRange("A5:J5").merge();
journal.getRange("K5:R5").merge();
journal.getRange("S5:W5").merge();
journal.getRange("A5:W5").values = [["План до сделки", "Учебное исполнение и параметры", "Разбор"]];
journal.getRange("A5:J5").values = [["План до сделки"]];
journal.getRange("K5:R5").values = [["Учебное исполнение и параметры"]];
journal.getRange("S5:W5").values = [["Результат и разбор"]];
journal.getRange("A5:W5").format = { fill: COLORS.blue, font: { bold: true, color: COLORS.white }, horizontalAlignment: "center" };
const journalHeaders = [
  "№", "Дата", "SECID", "Направление", "Учебная гипотеза", "Плановый вход",
  "Плановый стоп", "Контрактов", "Риск-бюджет", "Планируемый убыток до стопа",
  "Плановые комиссии", "Плановое проскальзывание", "Фактический вход",
  "Фактический выход", "Фактические комиссии", "Фактическое проскальзывание",
  "MINSTEP", "STEPPRICE", "Валовый результат", "Результат после издержек",
  "Риск-бюджет превышен?", "Качество решения", "Ошибка / следующий шаг",
];
journal.getRange("A6:W6").values = [journalHeaders];
header(journal.getRange("A6:W6"));
for (let row = 7; row <= 106; row += 1) {
  journal.getRange(`A${row}`).formulas = [[`=IF(B${row}=\"\",\"\",ROW()-6)`]];
  journal.getRange(`J${row}`).formulas = [[
    `=IF(OR(F${row}=\"\",G${row}=\"\",H${row}=\"\",Q${row}=\"\",R${row}=\"\"),\"\",ABS(F${row}-G${row})/Q${row}*R${row}*H${row}+K${row}+L${row})`,
  ]];
  journal.getRange(`Q${row}`).formulas = [[`=IF(C${row}=\"\",\"\",'Контракт'!$B$13)`]];
  journal.getRange(`R${row}`).formulas = [[`=IF(C${row}=\"\",\"\",'Контракт'!$B$14)`]];
  journal.getRange(`S${row}`).formulas = [[
    `=IF(OR(D${row}=\"\",M${row}=\"\",N${row}=\"\",H${row}=\"\",Q${row}=\"\",R${row}=\"\"),\"\",IF(D${row}=\"Лонг\",1,-1)*(N${row}-M${row})/Q${row}*R${row}*H${row})`,
  ]];
  journal.getRange(`T${row}`).formulas = [[`=IF(S${row}=\"\",\"\",S${row}-O${row}-P${row})`]];
  journal.getRange(`U${row}`).formulas = [[`=IF(T${row}=\"\",\"\",IF(MAX(0,-T${row})>I${row},\"Да\",\"Нет\"))`]];
}
input(journal.getRange("B7:I106"));
input(journal.getRange("K7:P106"));
input(journal.getRange("V7:W106"));
formula(journal.getRange("A7:A106"));
formula(journal.getRange("J7:J106"));
formula(journal.getRange("S7:U106"));
journal.getRange("Q7:R106").format = {
  fill: "#F3FBF4",
  font: { color: "#008000" },
  borders: { preset: "outside", style: "thin", color: COLORS.border },
};
journal.getRange("B7:B106").format.numberFormat = "yyyy-mm-dd";
journal.getRange("F7:T106").format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
journal.getRange("D7:D106").dataValidation = { rule: { type: "list", values: ["Лонг", "Шорт"] } };
journal.getRange("V7:V106").dataValidation = { rule: { type: "list", values: ["По процессу", "С отклонением", "Не по процессу"] } };
journal.getRange("U7:U106").conditionalFormats.add("containsText", {
  text: "Да",
  format: { fill: COLORS.lightRed, font: { color: COLORS.red, bold: true } },
});
journal.freezePanes.freezeRows(6);
journal.freezePanes.freezeColumns(4);
const journalTable = journal.tables.add("A6:W106", true, "PaperTradesTable");
journalTable.showBandedRows = true;
journalTable.showFilterButton = true;

// Прогресс
title(progress, "A1:F2", "Прогресс и интервальное повторение");
progress.getRange("A3:F3").merge();
progress.getRange("A3:F3").values = [[
  "Статус «самостоятельно» требует решения без подсказки; «перенос подтверждён» — новой проверки после задержки или в изменённом контексте.",
]];
note(progress.getRange("A3:F3"));
progress.getRange("A5:F5").values = [["Навык", "Статус", "Свидетельство", "Последняя проверка", "Следующая проверка", "Персональная ловушка"]];
header(progress.getRange("A5:F5"));
const progressRows = [
  ["Механика фьючерса", "требует диагностики", "Прошлое знакомство не подтверждает устойчивый навык", null, null, null],
  ["Плечо и база процента", "практика с подсказками", "После разбора решён близкий пример; нужен отсроченный контроль", new Date("2026-08-31T00:00:00"), null, "Не применять плечо дважды"],
  ["Паспорт контракта и единицы", "требует диагностики", "Нет свежего самостоятельного задания", null, null, null],
  ["P&L двумя способами", "практика с подсказками", "Направление и один путь знакомы; второй способ не закреплён", new Date("2026-08-31T00:00:00"), null, "Подписывать базу процента и единицы"],
  ["ГО, номинал, капитал и риск-бюджет", "не начато", "Расчёт числа контрактов по риск-бюджету ещё не проходил", null, null, "ГО не равно цене или максимуму убытка"],
  ["Стоп, исполнение и издержки", "не начато", "Не проверено", null, null, "Стоп не гарантирует цену исполнения"],
  ["Серия бумажных решений", "не начато", "Журнал пока пуст", null, null, null],
  ["Статистика и ограничения выборки", "требует диагностики", "Предметный перенос не проверен", null, null, null],
  ["Python для анализа", "не начато", "Практический опыт не подтверждён", null, null, "Вводить постепенно на знакомых данных"],
];
progress.getRange("A6:F14").values = progressRows;
input(progress.getRange("B6:F14"));
progress.getRange("D6:E14").format.numberFormat = "yyyy-mm-dd";
progress.getRange("B6:B14").dataValidation = {
  rule: {
    type: "list",
    values: ["не начато", "введено", "практика с подсказками", "самостоятельно", "перенос подтверждён", "нужно повторить", "требует диагностики"],
  },
};
progress.getRange("B6:B14").conditionalFormats.add("containsText", {
  text: "перенос подтверждён",
  format: { fill: COLORS.lightGreen, font: { color: COLORS.green, bold: true } },
});

// Источники
title(sources, "A1:F2", "Реестр источников и паспорт среза");
sources.getRange("A4:F4").values = [["Источник", "URL", "Роль", "Проверено", "Свежесть", "Ограничение"]];
header(sources.getRange("A4:F4"));
sources.getRange("A5:F12").values = [
  ["MOEX: срочный рынок", "https://www.moex.com/ru/derivatives", "Спецификации и документы", new Date("2026-09-01T00:00:00"), "Перед текущей задачей", "Проверять точную серию"],
  ["MOEX: ГО", "https://www.moex.com/ru/derivatives/go_futures.aspx", "Текущее гарантийное обеспечение", new Date("2026-09-01T00:00:00"), "В день использования", "ГО меняется"],
  ["MOEX ISS", "https://www.moex.com/a2193", "Карточки, снимки и свечи", new Date("2026-09-01T00:00:00"), "Каждый ответ с timestamp", "Данные могут быть отложены"],
  ["НКЦ", "https://www.nationalclearingcentre.ru/", "Клиринг и риск-параметры", new Date("2026-09-01T00:00:00"), "Действующая редакция", "Проверять область документа"],
  ["Банк России", "https://www.cbr.ru/s/258d", "Тестирование инвесторов", new Date("2026-09-01T00:00:00"), "Перед правовым выводом", "Проверять редакцию"],
  ["MOEX School", "https://school.moex.com/courses/srochnyj-rynok-klyuch-k-finansovoj-nezavisimosti", "Русскоязычное введение", new Date("2026-09-01T00:00:00"), "Периодически", "Не заменяет спецификацию"],
  ["CME Education", "https://www.cmegroup.com/education", "Механика фьючерсов", new Date("2026-09-01T00:00:00"), "Периодически", "Не переносить правила CME на MOEX"],
  ["MOEX AlgoPack", "https://moexalgo.github.io/docs/api/", "Будущий этап анализа данных", new Date("2026-09-01T00:00:00"), "Перед подключением", "Не нужен на базовом этапе"],
];
sources.getRange("D5:D12").format.numberFormat = "yyyy-mm-dd";
sources.getRange("A5:F12").format.wrapText = true;
sources.getRange("A5:F12").format.borders = { preset: "inside", style: "thin", color: COLORS.border };
section(sources, "A14:F14", "Паспорт использованного текущего значения");
sources.getRange("A15:F15").values = [["Тезис", "SECID", "Поле MOEX", "Значение и единица", "Источник URL / получено", "Статус и ограничения"]];
header(sources.getRange("A15:F15"));
input(sources.getRange("A16:F25"));

// Проверки
title(checks, "A1:F2", "Проверки модели");
checks.getRange("A4:F4").values = [["Проверка", "Фактическое", "Ожидаемое", "Разница", "Статус", "Что исправить"]];
header(checks.getRange("A4:F4"));
checks.getRange("A5:F10").values = [
  ["MINSTEP заполнен и положителен", null, ">0", null, null, "Заполнить Контракт!B13 из карточки MOEX"],
  ["STEPPRICE заполнен и положителен", null, ">0", null, null, "Заполнить Контракт!B14 и проверить спецификацию"],
  ["ГО заполнено и положительно", null, ">0", null, null, "Заполнить Контракт!B16; не трактовать как максимум убытка"],
  ["Строк с плановым убытком выше риск-бюджета", null, "0", null, null, "Пересмотреть число контрактов или план"],
  ["Строк с превышением риск-бюджета по факту", null, "0", null, null, "Разобрать исполнение и риск-процесс"],
  ["Источник текущего контракта указан", null, "Да", null, null, "Добавить прямой URL MOEX ISS"],
];
checks.getRange("B5:B10").formulas = [
  ["='Контракт'!$B$13"],
  ["='Контракт'!$B$14"],
  ["='Контракт'!$B$16"],
  ["=SUMPRODUCT(--('Журнал сделок'!$B$7:$B$106<>\"\"),--('Журнал сделок'!$J$7:$J$106>'Журнал сделок'!$I$7:$I$106))"],
  ["=COUNTIF('Журнал сделок'!$U$7:$U$106,\"Да\")"],
  ["=IF('Контракт'!$B$8<>\"\",\"Да\",\"Нет\")"],
];
checks.getRange("E5:E10").formulas = [
  ["=IF(B5>0,\"OK\",\"Проверить\")"],
  ["=IF(B6>0,\"OK\",\"Проверить\")"],
  ["=IF(B7>0,\"OK\",\"Проверить\")"],
  ["=IF(B8=0,\"OK\",\"Проверить\")"],
  ["=IF(B9=0,\"OK\",\"Проверить\")"],
  ["=IF(B10=\"Да\",\"OK\",\"Проверить\")"],
];
formula(checks.getRange("B5:B10"));
formula(checks.getRange("E5:E10"));
checks.getRange("B5:E10").format.horizontalAlignment = "center";
checks.getRange("A5:F10").format.borders = { preset: "inside", style: "thin", color: COLORS.border };
checks.getRange("E5:E10").conditionalFormats.add("containsText", {
  text: "OK",
  format: { fill: COLORS.lightGreen, font: { color: COLORS.green, bold: true } },
});
checks.getRange("E5:E10").conditionalFormats.add("containsText", {
  text: "Проверить",
  format: { fill: COLORS.lightRed, font: { color: COLORS.red, bold: true } },
});
checks.getRange("A12").values = [["Общий статус"]];
checks.getRange("B12").formulas = [["=IF(COUNTIF(E5:E10,\"Проверить\")=0,\"Модель готова\",\"Нужно заполнить и проверить\")"]];
checks.getRange("A12:B12").format = {
  fill: COLORS.lightAmber,
  font: { bold: true, color: COLORS.amber },
  borders: { preset: "outside", style: "medium", color: COLORS.amber },
};

// Shared sizing and visual hygiene.
overview.getRange("A1:H39").format.font.name = "Aptos";
contract.getRange("A1:D25").format.font.name = "Aptos";
calculator.getRange("A1:D35").format.font.name = "Aptos";
journal.getRange("A1:W106").format.font.name = "Aptos";
progress.getRange("A1:F14").format.font.name = "Aptos";
sources.getRange("A1:F25").format.font.name = "Aptos";
checks.getRange("A1:F12").format.font.name = "Aptos";

overview.getRange("A1:H39").format.rowHeight = 22;
overview.getRange("A3:H3").format.rowHeight = 42;
overview.getRange("A10:H13").format.rowHeight = 36;
overview.getRange("A16:H22").format.rowHeight = 32;
overview.getRange("A:A").format.columnWidth = 14;
overview.getRange("B:B").format.columnWidth = 15;
overview.getRange("C:C").format.columnWidth = 20;
overview.getRange("D:D").format.columnWidth = 15;
overview.getRange("E:E").format.columnWidth = 20;
overview.getRange("F:F").format.columnWidth = 15;
overview.getRange("G:G").format.columnWidth = 18;
overview.getRange("H:H").format.columnWidth = 15;

contract.getRange("A:A").format.columnWidth = 29;
contract.getRange("B:B").format.columnWidth = 28;
contract.getRange("C:C").format.columnWidth = 21;
contract.getRange("D:D").format.columnWidth = 48;
contract.getRange("A3:D3").format.rowHeight = 42;
contract.getRange("A6:D25").format.rowHeight = 32;

calculator.getRange("A:A").format.columnWidth = 36;
calculator.getRange("B:B").format.columnWidth = 28;
calculator.getRange("C:C").format.columnWidth = 18;
calculator.getRange("D:D").format.columnWidth = 48;
calculator.getRange("A3:D3").format.rowHeight = 42;
calculator.getRange("A6:D35").format.rowHeight = 30;
calculator.getRange("A6:D35").format.wrapText = true;

journal.getRange("A:A").format.columnWidth = 7;
journal.getRange("B:D").format.columnWidth = 14;
journal.getRange("E:E").format.columnWidth = 32;
journal.getRange("F:U").format.columnWidth = 17;
journal.getRange("V:V").format.columnWidth = 20;
journal.getRange("W:W").format.columnWidth = 34;
journal.getRange("A3:W3").format.rowHeight = 42;
journal.getRange("A5:W6").format.rowHeight = 38;
journal.getRange("A6:W106").format.wrapText = true;

progress.getRange("A:A").format.columnWidth = 34;
progress.getRange("B:B").format.columnWidth = 28;
progress.getRange("C:C").format.columnWidth = 56;
progress.getRange("D:E").format.columnWidth = 19;
progress.getRange("F:F").format.columnWidth = 40;
progress.getRange("A3:F3").format.rowHeight = 42;
progress.getRange("A5:F14").format.rowHeight = 36;
progress.getRange("A5:F14").format.wrapText = true;

sources.getRange("A:A").format.columnWidth = 24;
sources.getRange("B:B").format.columnWidth = 54;
sources.getRange("C:C").format.columnWidth = 32;
sources.getRange("D:D").format.columnWidth = 25;
sources.getRange("E:E").format.columnWidth = 32;
sources.getRange("F:F").format.columnWidth = 30;
sources.getRange("A15:F15").format.rowHeight = 52;
sources.getRange("A4:F25").format.rowHeight = 42;
sources.getRange("A4:F25").format.wrapText = true;

checks.getRange("A:A").format.columnWidth = 42;
checks.getRange("B:E").format.columnWidth = 18;
checks.getRange("F:F").format.columnWidth = 52;
checks.getRange("A4:F10").format.rowHeight = 36;
checks.getRange("A4:F12").format.wrapText = true;

overview.freezePanes.freezeRows(3);
contract.freezePanes.freezeRows(5);
calculator.freezePanes.freezeRows(5);
progress.freezePanes.freezeRows(5);
sources.freezePanes.freezeRows(4);
checks.freezePanes.freezeRows(4);

const overviewInspect = await workbook.inspect({
  kind: "table",
  range: "Обзор!A1:H22",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 10,
});
console.log(overviewInspect.ndjson);

const calculatorInspect = await workbook.inspect({
  kind: "table",
  range: "Калькулятор!A17:D35",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 6,
});
console.log(calculatorInspect.ndjson);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errorScan.ndjson);

for (const sheetName of ["Обзор", "Контракт", "Калькулятор", "Журнал сделок", "Прогресс", "Источники", "Проверки"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const safeName = sheetName.replaceAll(" ", "-");
  await fs.writeFile(path.join(outputDir, `preview-${safeName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "futures-learning-lab.xlsx"));
