const ExcelJS = require('exceljs');
const path = require('path');

module.exports = async function generateExcel() {
    const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');
    const outputPath = path.join(__dirname, 'public/templates/Golf_Tours_Generated.xlsx');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.getWorksheet(1);

    // Example data population (replace with your finalJson)
    worksheet.getCell('B2').value = 'Sample Golf Course';
    worksheet.getCell('C2').value = 'Sample Hotel';
    worksheet.getCell('D2').value = 'Sample Transport';

    await workbook.xlsx.writeFile(outputPath);
    console.log('Excel generated at:', outputPath);
};
