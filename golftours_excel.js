const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  // Pick correct template
  const templateFileName =
    data.currency_hint === '£'
      ? 'Golf_Tours_Template_Scotland.xlsx'
      : 'Golf_Tours_Template_Ireland.xlsx';
  const templatePath = path.join(__dirname, 'public/templates', templateFileName);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Determine Excel currency format
  const currencySymbol = data.currency_hint || '€';
  const currencyFormat = `${currencySymbol}#,##0.00`;

  // Basic info
  sheet.getCell('K5').value = `${data.lead_name} Group`;
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  // Row mapping
  const dayCellMap = [
    { date: 15, day_of_week: 16, hotel: 15, golf: 16, transport: 17 },
    { date: 20, day_of_week: 21, hotel: 20, golf: 21, transport: 22 },
    { date: 25, day_of_week: 26, hotel: 25, golf: 26, transport: 27 },
    { date: 30, day_of_week: 31, hotel: 30, golf: 31, transport: 32 },
    { date: 35, day_of_week: 36, hotel: 35, golf: 36, transport: 37 },
    { date: 40, day_of_week: 41, hotel: 40, golf: 41, transport: 42 },
    { date: 45, day_of_week: 46, hotel: 45, golf: 46, transport: 47 },
    { date: 50, day_of_week: 51, hotel: 50, golf: 51, transport: 52 },
    { date: 55, day_of_week: 56, hotel: 55, golf: 56, transport: 57 },
    { date: 60, day_of_week: 61, hotel: 60, golf: 61, transport: 62 },
    { date: 65, day_of_week: 66, hotel: 65, golf: 66, transport: 67 },
    { date: 70, day_of_week: 71, hotel: 70, golf: 71, transport: 72 },
  ];

  const itineraryDays = Object.keys(data.itinerary);

  itineraryDays.forEach((dayKey, index) => {
    const dayData = data.itinerary[dayKey];
    const map = dayCellMap[index];
    if (!map) return;

    const cleanNumber = (val) =>
      Number(String(val).replace(/[^\d.-]/g, '')) || 0;

    // Date & weekday
    sheet.getCell(`A${map.date}`).value = dayData.date;
    if (dayData.day_of_week) {
      sheet.getCell(`A${map.day_of_week}`).value = dayData.day_of_week;
    }

    // Hotel
    if (dayData.hotel_stay?.length) {
      const hotel = dayData.hotel_stay[0];
      sheet.getCell(`B${map.hotel}`).value = hotel.hotel;

      const sharing = cleanNumber(hotel.Hotel_Sharing);
      const single = cleanNumber(hotel.Hotel_Single);

      sheet.getCell(`C${map.hotel}`).value = sharing;
      sheet.getCell(`C${map.hotel}`).numFmt = currencyFormat;

      sheet.getCell(`D${map.hotel}`).value = single;
      sheet.getCell(`D${map.hotel}`).numFmt = currencyFormat;
    }

    // Golf
    if (dayData.Golf_round?.length) {
      const golf = dayData.Golf_round[0];
      sheet.getCell(`B${map.golf}`).value = golf.course;

      const golfRate = cleanNumber(golf.Golf);
      sheet.getCell(`E${map.golf}`).value = golfRate;
      sheet.getCell(`E${map.golf}`).numFmt = currencyFormat;
    }

    // Transport
    if (dayData.transport?.length) {
      const transport = dayData.transport[0];
      sheet.getCell(`B${map.transport}`).value = transport.transport_type;

      const rate = cleanNumber(transport.rate_per_person);
      sheet.getCell(`F${map.transport}`).value = rate;
      sheet.getCell(`F${map.transport}`).numFmt = currencyFormat;
    }
  });

  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcelWithDynamicItinerary };
