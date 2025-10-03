const ExcelJS = require('exceljs');

async function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Quotation Sheet');

  // Example headers
  sheet.addRow(['Team Member', 'Lead Name', 'Golfers', 'Non-Golfers']);
  sheet.addRow([data.team_member, data.lead_name, data.golfers, data.non_golfers]);

  sheet.addRow([]);
  sheet.addRow(['Itinerary']);

  // ✅ Loop through itinerary days
  Object.keys(data.itinerary).forEach((dayKey) => {
    const day = data.itinerary[dayKey];
    sheet.addRow([day.date]);

    // Golf rounds
    if (day.Golf_round) {
      day.Golf_round.forEach((g) => {
        sheet.addRow(['Golf Course', g.course, g.Golf, g.golf_hint || '']);
      });
    }

    // Hotel stays
    if (day.hotel_stay) {
      day.hotel_stay.forEach((h) => {
        sheet.addRow(['Hotel', h.hotel, h.Hotel_Single, h.Hotel_Sharing]);
      });
    }

    // Transport
    if (day.transport) {
      day.transport.forEach((t) => {
        sheet.addRow(['Transport', t.transport_type, t.total_people, t.rate_per_person]);
      });
    }

    sheet.addRow([]);
  });

  // Return Excel buffer
  return workbook.xlsx.writeBuffer();
}

// ✅ Correct CommonJS export
module.exports = { generateExcel };
