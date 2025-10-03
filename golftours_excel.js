const ExcelJS = require('exceljs');
const path = require('path');

module.exports = async function generateExcel() {
  try {
    // Path to template and output
    const templatePath = path.join(__dirname, 'public', 'templates', 'Golf_Tours_Template.xlsx');
    const outputPath = path.join(__dirname, 'public', 'templates', 'Golf_Tours_Generated.xlsx');

    // Load workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.getWorksheet(1);

    // Example: fill some sample data
    worksheet.getCell('B2').value = 'Sample Golf Course';
    worksheet.getCell('C2').value = 'Sample Hotel';
    worksheet.getCell('D2').value = 'Sample Transport';
    worksheet.getCell('E2').value = 100;

    // Save workbook
    await workbook.xlsx.writeFile(outputPath);

    console.log('Excel generated at:', outputPath);
    return outputPath; // Optional: return path if needed
  } catch (err) {
    console.error('Excel generation error:', err);
    throw err; // Important: rethrow so Express catches it
  }
};
