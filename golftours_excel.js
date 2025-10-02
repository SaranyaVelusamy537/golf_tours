const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Path to the uploaded template on Render disk
const templatePath = '/mnt/data/golf_quote_template.xlsx';
const outputPath = '/mnt/data/golf_quote_generated.xlsx';

// Load workbook
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(templatePath);

// Get the first worksheet (adjust if multiple sheets)
const worksheet = workbook.getWorksheet(1);

// Sample data (replace with your JSON)
const data = finalJson[0]; // finalJson is the output from previous script

// Loop through days to populate daily itinerary
Object.keys(data.itinerary).forEach((dayKey, index) => {
  const day = data.itinerary[dayKey];
  
  // Example: populate golf, hotel, transport
  worksheet.getCell(`B${index + 2}`).value = day.Golf_round[0].course;
  worksheet.getCell(`C${index + 2}`).value = day.hotel_stay[0].hotel;
  worksheet.getCell(`D${index + 2}`).value = day.transport[0].transport_type;
  worksheet.getCell(`E${index + 2}`).value = day.day_total[0].Combined_Single;
  worksheet.getCell(`F${index + 2}`).value = day.day_total[0].Combined_Sharing;
});

// Populate trip totals
worksheet.getCell('B20').value = data.trip_total.total_golf;
worksheet.getCell('B21').value = data.trip_total.total_hotel_single;
worksheet.getCell('B22').value = data.trip_total.total_hotel_sharing;
worksheet.getCell('B23').value = data.trip_total.total_transportation;

// Populate golfer margins
worksheet.getCell('D20').value = data.margin.golfer_margins.total_fit_rate_per_single;
worksheet.getCell('D21').value = data.margin.golfer_margins.total_fit_rate_per_sharing;

// Populate non-golfer margins
worksheet.getCell('E20').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;
worksheet.getCell('E21').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;

// Populate total tour margin
worksheet.getCell('F20').value = data.total_tour_margin.margin_40_total;
worksheet.getCell('F21').value = data.total_tour_margin.margin_35_total;

// Save populated Excel
await workbook.xlsx.writeFile(outputPath);

console.log('Excel generated at:', outputPath);
