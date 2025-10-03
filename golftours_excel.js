const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  // Load the pre-made template
  const path = require('path');

  const templatePath = path.join(__dirname, '..', 'public', 'templates', 'Golf_Tours_Template.xlsx');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Fill basic info
  sheet.getCell('K5').value = data.lead_name + ' Group'; // Lead Name
  sheet.getCell('L5').value = data.team_member;          // Team Member
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  // Fill FIT rates
  sheet.getCell('I12').value = data.margin.golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('J12').value = data.margin.golfer_margins.total_fit_rate_per_single;
  sheet.getCell('K12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;
  sheet.getCell('L12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;

  // Fill itinerary
  const dayStartRow = 15; // Row where Day 1 starts in template
  const dayRowIncrement = 3; // Each day uses 3 rows (Date/Hotel, Golf, Transport)
  const itineraryDays = Object.keys(data.itinerary);

  itineraryDays.forEach((dayKey, index) => {
    const dayData = data.itinerary[dayKey];
    const currentRow = dayStartRow + index * dayRowIncrement;

    // Row 1: Date + Hotel
    sheet.getCell(`A${currentRow}`).value = dayData.date; // Date
    if (dayData.hotel_stay && dayData.hotel_stay.length > 0) {
      const hotel = dayData.hotel_stay[0];
      sheet.getCell(`B${currentRow}`).value = hotel.hotel;           // Hotel Name
      sheet.getCell(`C${currentRow}`).value = hotel.Hotel_Sharing;   // Hotel Sharing
      sheet.getCell(`D${currentRow}`).value = hotel.Hotel_Single;    // Hotel Single
    }

    // Row 2: Day of Week + Golf
    if (dayData.Golf_round && dayData.Golf_round.length > 0) {
      const golf = dayData.Golf_round[0];
      sheet.getCell(`B${currentRow + 1}`).value = golf.course; // Golf Club Name
      sheet.getCell(`E${currentRow + 1}`).value = golf.Golf;   // Golf rate/value
    }

    // Row 3: Transport
    if (dayData.transport && dayData.transport.length > 0) {
      const transport = dayData.transport[0];
      sheet.getCell(`B${currentRow + 2}`).value = transport.transport_type; // Transport type
      sheet.getCell(`F${currentRow + 2}`).value = transport.rate_per_person; // Transport rate
    }
  });

  // Return Excel buffer
  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcelWithDynamicItinerary };
