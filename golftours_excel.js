const ExcelJS = require('exceljs');
const axios = require('axios');

async function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();

  // Load template from GitHub
  const response = await axios.get('https://raw.githubusercontent.com/SaranyaVelusamy537/golf_tours/main/public/templates/Golf_Tours_Template.xlsx', {
    responseType: 'arraybuffer',
  });

  await workbook.xlsx.load(response.data);

  const sheet = workbook.getWorksheet('Quotation Sheet'); // Use your template sheet name

  // Example: Inject calculated rates
  sheet.getCell('B2').value = data.team_member;
  sheet.getCell('B3').value = data.lead_name;
  sheet.getCell('C2').value = data.golfers;
  sheet.getCell('D2').value = data.non_golfers;

  // Inject itinerary dynamically
  let rowPointer = 6; // Example starting row
  Object.keys(data.itinerary).forEach((dayKey) => {
    const day = data.itinerary[dayKey];
    sheet.getCell(`A${rowPointer}`).value = day.date;

    if (day.Golf_round) {
      day.Golf_round.forEach((g) => {
        rowPointer++;
        sheet.getCell(`A${rowPointer}`).value = 'Golf Course';
        sheet.getCell(`B${rowPointer}`).value = g.course;
        sheet.getCell(`C${rowPointer}`).value = g.Golf;
      });
    }

    if (day.hotel_stay) {
      day.hotel_stay.forEach((h) => {
        rowPointer++;
        sheet.getCell(`A${rowPointer}`).value = 'Hotel';
        sheet.getCell(`B${rowPointer}`).value = h.hotel;
        sheet.getCell(`C${rowPointer}`).value = h.Hotel_Single;
        sheet.getCell(`D${rowPointer}`).value = h.Hotel_Sharing;
      });
    }

    if (day.transport) {
      day.transport.forEach((t) => {
        rowPointer++;
        sheet.getCell(`A${rowPointer}`).value = 'Transport';
        sheet.getCell(`B${rowPointer}`).value = t.transport_type;
        sheet.getCell(`C${rowPointer}`).value = t.total_people;
        sheet.getCell(`D${rowPointer}`).value = t.rate_per_person;
      });
    }

    rowPointer++;
  });

  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcel };
