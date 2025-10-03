const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcel(data) {
  const templatePath = path.join(__dirname, 'Golf_Tours_Template.xlsx'); // ensure the template is in the same folder
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Fill top-level info
  sheet.getCell('K5').value = `${data.lead_name} Group`; // Client Lead Name
  sheet.getCell('L5').value = ''; // optional if you have more details
  sheet.getCell('M5').value = ''; // optional if you have more details
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('J16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;
  sheet.getCell('L16').value = data.non_golfers;

  // Fill FIT rates
  sheet.getCell('I12').value = data.margin.golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('J12').value = data.margin.golfer_margins.total_fit_rate_per_single;
  sheet.getCell('K12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;
  sheet.getCell('L12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;

  // Map itinerary days
  const dayRowStartMap = {
    day1: 15,
    day2: 19,
    day3: 23,
    day4: 27,
    day5: 31,
    day6: 35,
    day7: 39,
    day8: 43,
    day9: 47,
    day10: 51,
    day11: 55,
    day12: 59
  };

  Object.keys(data.itinerary).forEach((dayKey) => {
    const day = data.itinerary[dayKey];
    const startRow = dayRowStartMap[dayKey];
    if (!startRow) return;

    // Date
    sheet.getCell(`A${startRow}`).value = day.date;

    // Hotel
    if (day.hotel_stay && day.hotel_stay[0]) {
      const hotel = day.hotel_stay[0];
      sheet.getCell(`B${startRow}`).value = hotel.hotel; // Hotel Name
      sheet.getCell(`C${startRow}`).value = hotel.Hotel_Sharing; // Sharing
      sheet.getCell(`D${startRow}`).value = hotel.Hotel_Single; // Single
    }

    // Golf
    if (day.Golf_round && day.Golf_round[0]) {
      const golf = day.Golf_round[0];
      sheet.getCell(`B${startRow + 1}`).value = golf.course; // Golf Club Name
      sheet.getCell(`E${startRow + 1}`).value = golf.Golf; // Golf rate
    }

    // Transport
    if (day.transport && day.transport[0]) {
      const t = day.transport[0];
      sheet.getCell(`B${startRow + 2}`).value = t.transport_type; // Transport type
      sheet.getCell(`F${startRow + 2}`).value = t.rate_per_person; // Rate per person
    }
  });

  // Return Excel buffer
  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcel };
