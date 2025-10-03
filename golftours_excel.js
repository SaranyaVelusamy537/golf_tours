const ExcelJS = require('exceljs');
const axios = require('axios');

async function generateExcelWithDynamicItinerary(data) {
  // Load template from GitHub
  const response = await axios.get(
    'https://raw.githubusercontent.com/SaranyaVelusamy537/golf_tours/main/public/templates/Golf_Tours_Template.xlsx',
    { responseType: 'arraybuffer' }
  );

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(response.data);

  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Fill general info
  sheet.getCell('K5').value = data.lead_name;
  sheet.getCell('L5').value = `${data.team_member} Group`;
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  // Map itinerary dynamically
  Object.keys(data.itinerary).forEach((dayKey, i) => {
    const day = data.itinerary[dayKey];

    // Find the row that contains "Day 1", "Day 2", etc.
    const dayHeader = `Day ${i + 1}`;
    let dayRow;
    sheet.eachRow((row, rowNumber) => {
      const firstCell = row.getCell(1).value;
      if (firstCell && firstCell.toString().includes(dayHeader)) {
        dayRow = rowNumber;
      }
    });

    if (!dayRow) {
      console.warn(`Cannot find row for ${dayHeader}`);
      return;
    }

    // Fill Date (assume first row after Day header)
    sheet.getCell(`A${dayRow + 1}`).value = day.date;

    // Fill Hotel (second row after Day header)
    if (day.hotel_stay && day.hotel_stay.length > 0) {
      const hotel = day.hotel_stay[0];
      sheet.getCell(`B${dayRow + 2}`).value = hotel.hotel; // Hotel Name
      sheet.getCell(`C${dayRow + 2}`).value = hotel.Hotel_Sharing;
      sheet.getCell(`D${dayRow + 2}`).value = hotel.Hotel_Single;
    }

    // Fill Golf (third row after Day header)
    if (day.Golf_round && day.Golf_round.length > 0) {
      const golf = day.Golf_round[0];
      sheet.getCell(`B${dayRow + 3}`).value = golf.course;
      sheet.getCell(`E${dayRow + 3}`).value = golf.Golf;
    }

    // Fill Transport (fourth row after Day header)
    if (day.transport && day.transport.length > 0) {
      const transport = day.transport[0];
      sheet.getCell(`B${dayRow + 4}`).value = transport.transport_type;
      sheet.getCell(`F${dayRow + 4}`).value = transport.rate_per_person;
    }
  });

  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcelWithDynamicItinerary };
