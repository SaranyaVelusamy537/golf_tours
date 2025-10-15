const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // ===== Basic info =====
  sheet.getCell('K5').value = (data.lead_name || 'Tour') + ' Group';
  sheet.getCell('I16').value = Number(data.golfers || 0);
  sheet.getCell('K16').value = Number(data.non_golfers || 0);

  // ===== 12-day cell map =====
  const dayCellMap = [
    { date: 15, day_of_week:16, hotel: 15, golf: 16, transport: 17 },
    { date: 20, day_of_week:21, hotel: 20, golf: 21, transport: 22 },
    { date: 25, day_of_week:26, hotel: 25, golf: 26, transport: 27 },
    { date: 30, day_of_week:31, hotel: 30, golf: 31, transport: 32 },
    { date: 35, day_of_week:36, hotel: 35, golf: 36, transport: 37 },
    { date: 40, day_of_week:41, hotel: 40, golf: 41, transport: 42 },
    { date: 45, day_of_week:46, hotel: 45, golf: 46, transport: 47 },
    { date: 50, day_of_week:51, hotel: 50, golf: 51, transport: 52 },
    { date: 55, day_of_week:56, hotel: 55, golf: 56, transport: 57 },
    { date: 60, day_of_week:61, hotel: 60, golf: 61, transport: 62 },
    { date: 65, day_of_week:66, hotel: 65, golf: 66, transport: 67 },
    { date: 70, day_of_week:71, hotel: 70, golf: 71, transport: 72 }
  ];

  const itineraryDays = Object.keys(data.itinerary || {});
  let lastCurrencySymbol = '€';

  // ===== Helpers =====
  const fmtFor = (sym) => `"${sym}"#,##0.00;[Red]\\-"${sym}"#,##0.00`;

  // Parse string like "£180", "€ 1,234.50", "1,234.5" → number; null/undefined/'' → null
  const parseAmount = (v) => {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return v;
    if (typeof v !== 'string') return Number(v) || null;
    // remove everything except digits, comma, dot, minus
    const cleaned = v.replace(/[^0-9,.\-]/g, '');
    if (!cleaned) return null;
    // handle simple comma thousands (assume dot decimal)
    const normalized = cleaned.replace(/,/g, '');
    const n = Number(normalized);
    return Number.isFinite(n) ? n : null;
  };

  const setMoney = (cell, value, currencySymbol) => {
    const num = parseAmount(value);
    cell.value = (num === null ? 0 : num);    // force a number so Excel shows 0.00 if empty
    cell.numFmt = fmtFor(currencySymbol);
  };

  // ===== Iterate through all 12 rows =====
  for (let i = 0; i < dayCellMap.length; i++) {
    const map = dayCellMap[i];
    const dayKey = itineraryDays[i];
    const dayData = dayKey ? (data.itinerary[dayKey] || {}) : {};

    // Currency symbol for this day (carry forward)
    const transport = dayData.transport?.[0];
    const currencySymbol = transport?.currency_hint || lastCurrencySymbol || '€';
    lastCurrencySymbol = currencySymbol;

    // Date & Day of Week
    sheet.getCell(`A${map.date}`).value = dayData.date || null;
    sheet.getCell(`A${map.day_of_week}`).value = dayData.day_of_week || null;

    // Hotel (text)
    const hotel = dayData.hotel_stay?.[0];
    sheet.getCell(`B${map.hotel}`).value = hotel?.hotel || null;

    // Hotel (money) — force numbers and format
    setMoney(sheet.getCell(`C${map.hotel}`), hotel?.Hotel_Sharing, currencySymbol);
    setMoney(sheet.getCell(`D${map.hotel}`), hotel?.Hotel_Single,  currencySymbol);

    // Golf (text + money)
    const golf = dayData.Golf_round?.[0];
    sheet.getCell(`B${map.golf}`).value = golf?.course || null;
    setMoney(sheet.getCell(`E${map.golf}`), golf?.Golf, currencySymbol);

    // Transport (text + money)
    sheet.getCell(`B${map.transport}`).value = transport?.transport_type || null;
    setMoney(sheet.getCell(`F${map.transport}`), transport?.rate_per_person, currencySymbol);
  }

  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcelWithDynamicItinerary };
