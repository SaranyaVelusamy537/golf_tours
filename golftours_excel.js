const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();

  // Load the pre-made template
  await workbook.xlsx.readFile(path.join(__dirname, 'Golf_Tours_Template.xlsx'));

  const sheet = workbook.getWorksheet('Quotation Sheet'); // change to your exact sheet name

  // -----------------------------
  // Inject general info
  // -----------------------------
  sheet.getCell('K5').value = data.lead_name;       // Client lead_name Group
  sheet.getCell('L5').value = data.team_member;

  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('J16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;
  sheet.getCell('L16').value = data.non_golfers;

  // -----------------------------
  // Inject FIT rates
  // -----------------------------
  sheet.getCell('I12').value = data.margin.golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('J12').value = data.margin.golfer_margins.total_fit_rate_per_single;
  sheet.getCell('K12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;
  sheet.getCell('L12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;

  // -----------------------------
  // Inject quoted rates per sharing
  // -----------------------------
  const sharingRates = data.margin.golfer_margins.quotated_rate_per_sharing;
  sheet.getCell('J21').value = sharingRates.margin_40;
  sheet.getCell('K21').value = sharingRates.margin_35;
  sheet.getCell('L21').value = sharingRates.margin_30;
  sheet.getCell('M21').value = sharingRates.margin_25;

  const singleRates = data.margin.golfer_margins.quotated_rate_per_single;
  sheet.getCell('J27').value = singleRates.margin_40;
  sheet.getCell('K27').value = singleRates.margin_35;
  sheet.getCell('L27').value = singleRates.margin_30;
  sheet.getCell('M27').value = singleRates.margin_25;

  // -----------------------------
  // Inject margin per golfer
  // -----------------------------
  const sharingMargin = data.margin.golfer_margins.margin_per_sharing;
  sheet.getCell('J22').value = sharingMargin.margin_40_difference;
  sheet.getCell('K22').value = sharingMargin.margin_35_difference;
  sheet.getCell('L22').value = sharingMargin.margin_30_difference;
  sheet.getCell('M22').value = sharingMargin.margin_25_difference;

  const singleMargin = data.margin.golfer_margins.margin_per_single;
  sheet.getCell('J28').value = singleMargin.margin_40_difference;
  sheet.getCell('K28').value = singleMargin.margin_35_difference;
  sheet.getCell('L28').value = singleMargin.margin_30_difference;
  sheet.getCell('M28').value = singleMargin.margin_25_difference;

  // -----------------------------
  // Inject total group margin
  // -----------------------------
  const sharingGroup = data.margin.golfer_margins.total_group_margin_sharing;
  sheet.getCell('J23').value = sharingGroup.margin_40_group;
  sheet.getCell('K23').value = sharingGroup.margin_35_group;
  sheet.getCell('L23').value = sharingGroup.margin_30_group;
  sheet.getCell('M23').value = sharingGroup.margin_25_group;

  const singleGroup = data.margin.golfer_margins.total_group_margin_single;
  sheet.getCell('J29').value = singleGroup.margin_40_group;
  sheet.getCell('K29').value = singleGroup.margin_35_group;
  sheet.getCell('L29').value = singleGroup.margin_30_group;
  sheet.getCell('M29').value = singleGroup.margin_25_group;

  // -----------------------------
  // Inject non-golfer margins
  // -----------------------------
  const nonSharingRates = data.margin.non_golfer_margins.quotated_rate_per_nongolfer_sharing;
  sheet.getCell('J33').value = nonSharingRates.margin_40;
  sheet.getCell('K33').value = nonSharingRates.margin_35;
  sheet.getCell('L33').value = nonSharingRates.margin_30;
  sheet.getCell('M33').value = nonSharingRates.margin_25;

  const nonSingleRates = data.margin.non_golfer_margins.quotated_rate_per_nongolfer_single;
  sheet.getCell('J39').value = nonSingleRates.margin_40;
  sheet.getCell('K39').value = nonSingleRates.margin_35;
  sheet.getCell('L39').value = nonSingleRates.margin_30;
  sheet.getCell('M39').value = nonSingleRates.margin_25;

  const nonSharingMargin = data.margin.non_golfer_margins.margin_per_nongolfer_sharing;
  sheet.getCell('J34').value = nonSharingMargin.margin_40_difference;
  sheet.getCell('K34').value = nonSharingMargin.margin_35_difference;
  sheet.getCell('L34').value = nonSharingMargin.margin_30_difference;
  sheet.getCell('M34').value = nonSharingMargin.margin_25_difference;

  const nonSingleMargin = data.margin.non_golfer_margins.margin_per_nongolfer_single;
  sheet.getCell('J40').value = nonSingleMargin.margin_40_difference;
  sheet.getCell('K40').value = nonSingleMargin.margin_35_difference;
  sheet.getCell('L40').value = nonSingleMargin.margin_30_difference;
  sheet.getCell('M40').value = nonSingleMargin.margin_25_difference;

  // Total group margins non-golfers
  const nonSharingGroup = data.margin.non_golfer_margins.total_group_margin_sharing;
  sheet.getCell('J35').value = nonSharingGroup.margin_40_group;
  sheet.getCell('K35').value = nonSharingGroup.margin_35_group;
  sheet.getCell('L35').value = nonSharingGroup.margin_30_group;
  sheet.getCell('M35').value = nonSharingGroup.margin_25_group;

  const nonSingleGroup = data.margin.non_golfer_margins.total_group_margin_single;
  sheet.getCell('J41').value = nonSingleGroup.margin_40_group;
  sheet.getCell('K41').value = nonSingleGroup.margin_35_group;
  sheet.getCell('L41').value = nonSingleGroup.margin_30_group;
  sheet.getCell('M41').value = nonSingleGroup.margin_25_group;

  // -----------------------------
  // Inject itinerary per day
  // -----------------------------
  Object.keys(data.itinerary).forEach((dayKey, idx) => {
    const day = data.itinerary[dayKey];
    const dayNum = idx + 1;

    // Map days to your template rows manually
    // Example Day 1 (adjust row numbers according to template)
    if (dayNum === 1) {
      sheet.getCell('A15').value = day.date;
      sheet.getCell('B16').value = day.Golf_round[0].course;
      sheet.getCell('E16').value = day.Golf_round[0].Golf;
      sheet.getCell('C15').value = day.hotel_stay[0].Hotel_Sharing;
      sheet.getCell('D15').value = day.hotel_stay[0].Hotel_Single;
      sheet.getCell('B17').value = day.transport[0].transport_type;
      sheet.getCell('F22').value = day.transport[0].rate_per_person;
    }

    // Repeat mapping for day2 -> day7
    // Adjust row numbers accordingly
  });

  // -----------------------------
  // Inject trip totals
  // -----------------------------
  const totals = data.trip_total;
  sheet.getCell('C12').value = totals.total_hotel_sharing;
  sheet.getCell('D12').value = totals.total_hotel_single;
  sheet.getCell('E12').value = totals.total_golf;
  sheet.getCell('F12').value = totals.total_transportation;

  // -----------------------------
  // Return Excel buffer
  // -----------------------------
  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcel };
