const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  // Load the template
  const workbook = new ExcelJS.Workbook();
  const templatePath = path.join(__dirname, 'Golf_Tours_Template.xlsx');
  await workbook.xlsx.readFile(templatePath);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  if (!sheet) throw new Error('Sheet "Quotation Sheet" not found');

  // Fill basic info
  sheet.getCell('K5').value = data.lead_name; // Client Lead Name
  sheet.getCell('L5').value = data.lead_name; // Adjust if needed
  sheet.getCell('M5').value = data.lead_name;

  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('J16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;
  sheet.getCell('L16').value = data.non_golfers;

  // Fill FIT rates
  sheet.getCell('I12').value = data.margin.golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('J12').value = data.margin.golfer_margins.total_fit_rate_per_single;
  sheet.getCell('K12').value = data.margin.non_golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('L12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;

  // Fill itinerary
  const dayStartRows = {
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
    const dayData = data.itinerary[dayKey];
    const startRow = dayStartRows[dayKey];

    if (!startRow) return;

    // Date
    sheet.getCell(`A${startRow}`).value = dayData.date;

    // Hotel
    if (dayData.hotel_stay && dayData.hotel_stay.length > 0) {
      const hotel = dayData.hotel_stay[0];
      sheet.getCell(`B${startRow}`).value = hotel.hotel;
      sheet.getCell(`C${startRow}`).value = hotel.Hotel_Sharing;
      sheet.getCell(`D${startRow}`).value = hotel.Hotel_Single;
    }

    /
