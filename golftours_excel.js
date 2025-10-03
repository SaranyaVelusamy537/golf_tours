const ExcelJS = require('exceljs');
const path = require('path');

module.exports = async function generateExcel() {
  try {
    console.log('Starting Excel generation...');

    const templatePath = path.join(__dirname, 'public', 'templates', 'Golf_Tours_Template.xlsx');
    const outputPath = path.join(__dirname, 'public', 'templates', 'Golf_Tours_Generated.xlsx');

    console.log('Template path:', templatePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath); // This is often the place errors occur

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) throw new Error('Worksheet not found in template');

    console.log('Worksheet loaded, filling data...');

    // Example data filling
    worksheet.getCell('B2').value = 'Sample Golf Course';
    worksheet.getCell('C2').value = 'Sample Hotel';
    worksheet.getCell('D2').value = 'Sample Transport';
    worksheet.getCell('E2').value = 100;

    await workbook.xlsx.writeFile(outputPath);

    console.log('Excel generated at:', outputPath);
    return outputPath;
  } catch (err) {
    console.error('Excel generation failed:', err);
    throw err; // Let server.js catch this
  }
};
