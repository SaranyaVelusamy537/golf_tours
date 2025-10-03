const ExcelJS = require('exceljs');
const axios = require('axios');

async function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();

  // ✅ Load template from GitHub
  const response = await axios.get(
    'https://raw.githubusercontent.com/SaranyaVelusamy537/golf_tours/main/public/templates/Golf_Tours_Template.xlsx',
    { responseType: 'arraybuffer' }
  );
  await workbook.xlsx.load(response.data);

  // Assume first worksheet
  const sheet = workbook.getWorksheet(1);

  // ----------------------
  // Inject general info
  // ----------------------
  sheet.getCell('K5').value = `Client ${data.lead_name} Group`;
  sheet.getCell('L5').value = data.team_member;
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  // ----------------------
  // FIT rates
  // ----------------------
  sheet.getCell('J12').value = data.margin.golfer_margins.total_fit_rate_per_single;
  sheet.getCell('I12').value = data.margin.golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('L12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;
  sheet.getCell('K12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;

  // ----------------------
  // Quoted rates per sharing / single
  // ----------------------
  const gm = data.margin.golfer_margins;
  const ngm = data.margin.non_golfer_margins;

  // Golfer sharing
  sheet.getCell('J21').value = gm.quotated_rate_per_sharing.margin_40;
  sheet.getCell('K21').value = gm.quotated_rate_per_sharing.margin_35;
  sheet.getCell('L21').value = gm.quotated_rate_per_sharing.margin_30;
  sheet.getCell('M21').value = gm.quotated_rate_per_sharing.margin_25;

  // Golfer single
  sheet.getCell('J27').value = gm.quotated_rate_per_single.margin_40;
  sheet.getCell('K27').value = gm.quotated_rate_per_single.margin_35;
  sheet.getCell('L27').value = gm.quotated_rate_per_single.margin_30;
  sheet.getCell('M27').value = gm.quotated_rate_per_single.margin_25;

  // Margin differences
  sheet.getCell('J22').value = gm.margin_per_sharing.margin_40_difference;
  sheet.getCell('K22').value = gm.margin_per_sharing.margin_35_difference;
  sheet.getCell('L22').value = gm.margin_per_sharing.margin_30_difference;
  sheet.getCell('M22').value = gm.margin_per_sharing.margin_25_difference;

  sheet.getCell('J28').value = gm.margin_per_single.margin_40_difference;
  sheet.getCell('K28').value = gm.margin_per_single.margin_35_difference;
  sheet.getCell('L28').value = gm.margin_per_single.margin_30_difference;
  sheet.getCell('M28').value = gm.margin_per_single.margin_25_difference;

  // Total group margin
  sheet.getCell('J23').value = gm.total_group_margin_sharing.margin_40_group;
  sheet.getCell('K23').value = gm.total_group_margin_sharing.margin_35_group;
  sheet.getCell('L23').value = gm.total_group_margin_sharing.margin_30_group;
  sheet.getCell('M23').value = gm.total_group_margin_sharing.margin_25_group;

  sheet.getCell('J29').value = gm.total_group_margin_single.margin_40_group;
  sheet.getCell('K29').value = gm.total_group_margin_single.margin_35_group;
  sheet.getCell('L29').value = gm.total_group_margin_single.margin_30_group;
  sheet.getCell('M29').value = gm.total_group_margin_single.margin_25_group;

  // ----------------------
  // Non-golfer quoted rates
  // ----------------------
  sheet.getCell('J33').value = ngm.quotated_rate_per_nongolfer_sharing.margin_40;
  sheet.getCell('K33').value = ngm.quotated_rate_per_nongolfer_sharing.margin_35;
  sheet.getCell('L33').value = ngm.quotated_rate_per_nongolfer_sharing.margin_30;
  sheet.getCell('M33').value = ngm.quotated_rate_per_nongolfer_sharing.margin_25;

  sheet.getCell('J39').value = ngm.quotated_rate_per_nongolfer_single.margin_40;
  sheet.getCell('K39').value = ngm.quotated_rate_per_nongolfer_single.margin_35;
  sheet.getCell('L39').value = ngm.quotated_rate_per_nongolfer_single.margin_30;
  sheet.getCell('M39').value = ngm.quotated_rate_per_nongolfer_single.margin_25;

  sheet.getCell('J34').value = ngm.margin_per_nongolfer_sharing.margin_40_difference;
  sheet.getCell('K34').value = ngm.margin_per_nongolfer_sharing.margin_35_difference;
  sheet.getCell('L34').value = ngm.margin_per_nongolfer_sharing.margin_30_difference;
  sheet.getCell('M34').value = ngm.margin_per_nongolfer_sharing.margin_25_difference;

  sheet.getCell('J40').value = ngm.margin_per_nongolfer_single.margin_40_difference;
  sheet.getCell('K40').value = ngm.margin_per_nongolfer_single.margin_35_difference;
  sheet.getCell('L40').value = ngm.margin_per_nongolfer_single.margin_30_difference;
  sheet.getCell('M40').value = ngm.margin_per_nongolfer_single.margin_25_difference;

  // ----------------------
  // Trip totals
  // ----------------------
  sheet.getCell('C12').value = data.trip_total.total_hotel_sharing;
  sheet.getCell('D12').value = data.trip_total.total_hotel_single;
  sheet.getCell('E12').value = data.trip_total.total_golf;
  sheet.getCell('F12').value = data.trip_total.total_transportation;

  // ----------------------
  // Itinerary
  // ----------------------
  Object.keys(data.itinerary).forEach((dayKey, index) => {
    const day = data.itinerary[dayKey];
    const rowOffset = 15 + index * 4; // Adjust row offsets based on template structure

    // Date
    sheet.getCell(`A${rowOffset}`).value = day.date;

    // Golf rounds
    if (day.Golf_round) {
      day.Golf_round.forEach((g, i) => {
        sheet.getCell(`B${rowOffset + i + 1}`).value = g.course;
        sheet.getCell(`E${rowOffset + i + 1}`).value = g.Golf;
        sheet.getCell(`F${rowOffset + i + 1}`).value = g.golf_hint || '';
      });
    }

    // Hotel
    if (day.hotel_stay) {
      day.hotel_stay.forEach((h, i) => {
        sheet.getCell(`B${rowOffset + i}`).value = h.hotel;
        sheet.getCell(`C${rowOffset + i}`).value = h.Hotel_Sharing;
        sheet.getCell(`D${rowOffset + i}`).value = h.Hotel_Single;
      });
    }

    // Transport
    if (day.transport) {
      day.transport.forEach((t, i) => {
        sheet.getCell(`B${rowOffset + i}`).value = t.transport_type;
        sheet.getCell(`F${rowOffset + i}`).value = t.rate_per_person;
      });
    }
  });

  // ----------------------
  // Return Excel buffer
  // ----------------------
  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcel };
