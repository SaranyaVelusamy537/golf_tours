const ExcelJS = require('exceljs');
const axios = require('axios');

async function generateExcel(data) {
  // 1️⃣ Fetch template from GitHub
  const response = await axios.get(
    'https://raw.githubusercontent.com/SaranyaVelusamy537/golf_tours/main/public/templates/Golf_Tours_Template.xlsx',
    { responseType: 'arraybuffer' }
  );

  // 2️⃣ Load workbook from buffer
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(response.data);

  // 3️⃣ Get the correct sheet
  const sheet = workbook.getWorksheet('Quotation Sheet');
  if (!sheet) throw new Error("Worksheet 'Quotation Sheet' not found");

  // 4️⃣ Inject values into specific cells
  sheet.getCell('K5').value = `Client ${data.lead_name} Group`; // Lead Name
  sheet.getCell('L5').value = data.lead_name;

  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('J16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;
  sheet.getCell('L16').value = data.non_golfers;

  // Fit rates example
  sheet.getCell('I12').value = data.margin.golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('J12').value = data.margin.golfer_margins.total_fit_rate_per_single;
  sheet.getCell('K12').value = data.margin.non_golfer_margins.total_fit_rate_per_sharing;
  sheet.getCell('L12').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;

  // Quoted rates per golfer (sharing)
  sheet.getCell('J21').value = data.margin.golfer_margins.quotated_rate_per_sharing.margin_40;
  sheet.getCell('K21').value = data.margin.golfer_margins.quotated_rate_per_sharing.margin_35;
  sheet.getCell('L21').value = data.margin.golfer_margins.quotated_rate_per_sharing.margin_30;
  sheet.getCell('M21').value = data.margin.golfer_margins.quotated_rate_per_sharing.margin_25;

  // Quoted rates per golfer (single)
  sheet.getCell('J27').value = data.margin.golfer_margins.quotated_rate_per_single.margin_40;
  sheet.getCell('K27').value = data.margin.golfer_margins.quotated_rate_per_single.margin_35;
  sheet.getCell('L27').value = data.margin.golfer_margins.quotated_rate_per_single.margin_30;
  sheet.getCell('M27').value = data.margin.golfer_margins.quotated_rate_per_single.margin_25;

  // … Repeat for all other cells you want to inject

  // 5️⃣ Return buffer for sending as response
  return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcel };
