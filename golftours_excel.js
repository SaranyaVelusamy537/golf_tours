const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');  

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Fill basic info
  sheet.getCell('K5').value = data.lead_name + ' Group';
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  // Fill itinerary using fixed row map
  const dayCellMap = [
    { date: 15, day_of_week:16, hotel: 15, golf: 16, transport: 17 }, // Day 1
    { date: 20, day_of_week:21, hotel: 20, golf: 21, transport: 22 }, // Day 2
    { date: 25, day_of_week:26, hotel: 25, golf: 26, transport: 27 }, // Day 3
    { date: 30, day_of_week:31, hotel: 30, golf: 31, transport: 32 }, // Day 4
    { date: 35, day_of_week:36, hotel: 35, golf: 36, transport: 37 }, // Day 5
    { date: 40, day_of_week:41, hotel: 40, golf: 41, transport: 42 }, // Day 6
    { date: 45, day_of_week:46, hotel: 45, golf: 46, transport: 47 }, // Day 7
    { date: 50, day_of_week:50, hotel: 50, golf: 51, transport: 52 }, // Day 8
    { date: 55, day_of_week:56, hotel: 55, golf: 56, transport: 57 }, // Day 9
    { date: 60, day_of_week:61, hotel: 60, golf: 61, transport: 62 }, // Day 10
    { date: 65, day_of_week:66, hotel: 65, golf: 66, transport: 67 }, // Day 11
    { date: 70, day_of_week:71, hotel: 70, golf: 71, transport: 72 }  // Day 12
  ];

  const itineraryDays = Object.keys(data.itinerary);

  // Determine currency symbol from first day's transport.currency_hint (fallback €)
  const firstDayTransport = itineraryDays.length && data.itinerary[itineraryDays[0]].transport?.[0];
  const currencySymbol = firstDayTransport?.currency_hint || '€';
  const currencyFormat = `"${currencySymbol}"#,##0.00;[Red]\-"${currencySymbol}"#,##0.00`;

  itineraryDays.forEach((dayKey, index) => {
    const dayData = data.itinerary[dayKey];
    const map = dayCellMap[index];
    if (!map) return;

    // Row 1: Date
    sheet.getCell(`A${map.date}`).value = dayData.date;

    // Row 1: Day of Week
    if (dayData.day_of_week) {
      sheet.getCell(`A${map.day_of_week}`).value = dayData.day_of_week;
    }

    // Hotel
    if (dayData.hotel_stay && dayData.hotel_stay.length > 0) {
      const hotel = dayData.hotel_stay[0];
      sheet.getCell(`B${map.hotel}`).value = hotel.hotel;
      sheet.getCell(`C${map.hotel}`).value = hotel.Hotel_Sharing;
      sheet.getCell(`D${map.hotel}`).value = hotel.Hotel_Single;

      // Apply currency format
      ['C','D'].forEach(col => {
        sheet.getCell(`${col}${map.hotel}`).numFmt = currencyFormat;
      });
    } else {
      // Apply currency format to empty cells
      ['C','D'].forEach(col => {
        sheet.getCell(`${col}${map.hotel}`).numFmt = currencyFormat;
      });
    }

    // Golf
    if (dayData.Golf_round && dayData.Golf_round.length > 0) {
      const golf = dayData.Golf_round[0];
      sheet.getCell(`B${map.golf}`).value = golf.course;
      sheet.getCell(`E${map.golf}`).value = golf.Golf;
      sheet.getCell(`E${map.golf}`).numFmt = currencyFormat;
    } else {
      sheet.getCell(`E${map.golf}`).numFmt = currencyFormat;
    }

    // Transport
    if (dayData.transport && dayData.transport.length > 0) {
      const transport = dayData.transport[0];
      sheet.getCell(`B${map.transport}`).value = transport.transport_type;
      sheet.getCell(`F${map.transport}`).value = transport.rate_per_person;
      sheet.getCell(`F${map.transport}`).numFmt = currencyFormat;
    } else {
      sheet.getCell(`F${map.transport}`).numFmt = currencyFormat;
    }
  });

  // Also ensure remaining day rows (Day 6-12) have currency format even if empty
  dayCellMap.slice(itineraryDays.length).forEach(map => {
    ['C','D'].forEach(col => sheet.getCell(`${col}${map.hotel}`).numFmt = currencyFormat);
    sheet.getCell(`E${map.golf}`).numFmt = currencyFormat;
    sheet.getCell(`F${map.transport}`).numFmt = currencyFormat;
  });

  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcelWithDynamicItinerary };
