// golftours_excel.js
const ExcelJS = require('exceljs');
const path = require('path');

// Paths
const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');
const outputPath = '/mnt/data/golf_quote_generated.xlsx';

/**
 * Generates a new Excel file from the template and provided data
 * @param {Array} finalJson - Array with your calculated rates and itinerary
 * @returns {Promise<string>} - Path to generated Excel file
 */
async function generateExcel(finalJson) {
  if (!finalJson || !Array.isArray(finalJson) || finalJson.length === 0) {
    throw new Error('Invalid finalJson data');
  }

  const data = finalJson[0];

  // Load workbook
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  // Get first worksheet (adjust if needed)
  const worksheet = workbook.getWorksheet(1);

  // Populate daily itinerary
  Object.keys(data.itinerary).forEach((dayKey, index) => {
    const day = data.itinerary[dayKey];

    // Adjust cell positions according to your template layout
    worksheet.getCell(`B${index + 2}`).value = day.Golf_round[0].course || '';
    worksheet.getCell(`C${index + 2}`).value = day.hotel_stay[0].hotel || '';
    worksheet.getCell(`D${index + 2}`).value = day.transport[0].transport_type || '';
    worksheet.getCell(`E${index + 2}`).value = day.day_total[0].Combined_Single || 0;
    worksheet.getCell(`F${index + 2}`).value = day.day_total[0].Combined_Sharing || 0;
  });

  // Populate trip totals
  worksheet.getCell('B20').value = data.trip_total.total_golf || 0;
  worksheet.getCell('B21').value = data.trip_total.total_hotel_single || 0;
  worksheet.getCell('B22').value = data.trip_total.total_hotel_sharing || 0;
  worksheet.getCell('B23').value = data.trip_total.total_transportation || 0;

  // Populate golfer margins
  worksheet.getCell('D20').value = data.margin.golfer_margins.total_fit_rate_per_single || 0;
  worksheet.getCell('D21').value = data.margin.golfer_margins.total_fit_rate_per_sharing || 0;

  // Populate non-golfer margins
  worksheet.getCell('E20').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single || 0;
  worksheet.getCell('E21').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing || 0;

  // Populate total tour margin
  worksheet.getCell('F20').value = data.total_tour_margin.margin_40_total || 0;
  worksheet.getCell('F21').value = data.total_tour_margin.margin_35_total || 0;

  // Save generated Excel
  await workbook.xlsx.writeFile(outputPath);
  console.log('Excel generated at:', outputPath);

  return outputPath;
}

// Export the function
module.exports = generateExcel;
