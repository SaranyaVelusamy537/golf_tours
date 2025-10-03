const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcel(finalJson) {
    const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');
    const outputPath = path.join(__dirname, 'public/templates/Golf_Tours_Generated.xlsx');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.getWorksheet(1);

    // Your existing population logic here
    // Example:
    Object.keys(finalJson.itinerary).forEach((dayKey, index) => {
        const day = finalJson.itinerary[dayKey];
        worksheet.getCell(`B${index + 2}`).value = day.Golf_round[0].course;
        worksheet.getCell(`C${index + 2}`).value = day.hotel_stay[0].hotel;
        worksheet.getCell(`D${index + 2}`).value = day.transport[0].transport_type;
        worksheet.getCell(`E${index + 2}`).value = day.day_total[0].Combined_Single;
        worksheet.getCell(`F${index + 2}`).value = day.day_total[0].Combined_Sharing;
    });

    // Populate totals and margins as before
    worksheet.getCell('B20').value = finalJson.trip_total.total_golf;
    worksheet.getCell('B21').value = finalJson.trip_total.total_hotel_single;
    worksheet.getCell('B22').value = finalJson.trip_total.total_hotel_sharing;
    worksheet.getCell('B23').value = finalJson.trip_total.total_transportation;

    worksheet.getCell('D20').value = finalJson.margin.golfer_margins.total_fit_rate_per_single;
    worksheet.getCell('D21').value = finalJson.margin.golfer_margins.total_fit_rate_per_sharing;
    worksheet.getCell('E20').value = finalJson.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;
    worksheet.getCell('E21').value = finalJson.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;
    worksheet.getCell('F20').value = finalJson.total_tour_margin.margin_40_total;
    worksheet.getCell('F21').value = finalJson.total_tour_margin.margin_35_total;

    await workbook.xlsx.writeFile(outputPath);
    console.log('Excel generated at:', outputPath);

    return outputPath;
}

module.exports = { generateExcel };
