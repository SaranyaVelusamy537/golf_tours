const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');  

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Fill basic info
  sheet.getCell('K5').value = data.lead_name + ' Group';
  sheet.getCell('L5').value = data.team_member;
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  // Fill itinerary using fixed row map
  const dayCellMap = [
    { date: 15, hotel: 15, golf: 16, transport: 17 },
    { date: 20, hotel: 20, golf: 21, transport: 22 },
    { date: 25, hotel: 25, golf: 26, transport: 27 },
    { date: 30, hotel: 30, golf: 31, transport: 32 },
    { date: 35, hotel: 35, golf: 36, transport: 37 },
    { date: 40, hotel: 40, golf: 41, transport: 42 },
    { date: 45, hotel: 45, golf: 46, transport: 47 },
    { date: 50, hotel: 50, golf: 51, transport: 52 },
    { date: 55, hotel: 55, golf: 56, transport: 57 },
    { date: 60, hotel: 60, golf: 61, transport: 62 },
    { date: 65, hotel: 65, golf: 66, transport: 67 },
    { date: 70, hotel: 70, golf: 71, transport: 72 },
  ];

  const itineraryDays = Object.keys(data.itinerary);

  itineraryDays.forEach((dayKey, index) => {
    const dayData = data.itinerary[dayKey];
    const map = dayCellMap[index];
    if (!map) return;

    // Row 1: Date + Hotel
    sheet.getCell(`A${map.date}`).value = dayData.date;
    if (dayData.hotel_stay && dayData.hotel_stay.length > 0) {
      const hotel = dayData.hotel_stay[0];
      sheet.getCell(`B${map.hotel}`).value = hotel.hotel;
      sheet.getCell(`C${map.hotel}`).value = hotel.Hotel_Sharing;
      sheet.getCell(`D${map.hotel}`).value = hotel.Hotel_Single;
    }

    // Row 2: Day of Week + Golf
    if (dayData.Golf_round && dayData.Golf_round.length > 0) {
      const golf = dayData.Golf_round[0];
      sheet.getCell(`B${map.golf}`).value = golf.course;
      sheet.getCell(`E${map.golf}`).value = golf.Golf;
    }

    // Row 3: Transport
    if (dayData.transport && dayData.transport.length > 0) {
      const transport = dayData.transport[0];
      sheet.getCell(`B${map.transport}`).value = transport.transport_type;
      sheet.getCell(`F${map.transport}`).value = transport.rate_per_person;
    }
  });

  // Generate buffer
  const buffer = await workbook.xlsx.writeBuffer();

  // Fix the download file name
  const groupName = data.lead_name + ' Group';
  const groupFilename = groupName.trim().toLowerCase().replace(/\s+/g, '_') + '.xlsx';

  // Return as buffer + filename (for HTTP response or n8n binary)
  return {
    buffer,
    fileName: groupFilename,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  };
}

module.exports = { generateExcelWithDynamicItinerary };
