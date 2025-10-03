const ExcelJS = require('exceljs');
const path = require('path');

module.exports = async function generateExcel(finalJson) {
    const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');
    const outputPath = path.join(__dirname, 'public/templates/Golf_Tours_Generated.xlsx');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.getWorksheet('Quotation Sheet'); // use your actual sheet name

    const data = finalJson[0];

    Object.keys(data.itinerary).forEach((dayKey, index) => {
        const day = data.itinerary[dayKey];
        worksheet.getCell(`B${index + 2}`).value = day.Golf_round[0].course;
        worksheet.getCell(`C${index + 2}`).value = day.hotel_stay[0].hotel;
        worksheet.getCell(`D${index + 2}`).value = day.transport[0].transport_type;
        worksheet.getCell(`E${index + 2}`).value = day.day_total[0].Combined_Single;
        worksheet.getCell(`F${index + 2}`).value = day.day_total[0].Combined_Sharing;
    });

    worksheet.getCell('B20').value = data.trip_total.total_golf;
    worksheet.getCell('B21').value = data.trip_total.total_hotel_single;
    worksheet.getCell('B22').value = data.trip_total.total_hotel_sharing;
    worksheet.getCell('B23').value = data.trip_total.total_transportation;

    worksheet.getCell('D20').value = data.margin.golfer_margins.total_fit_rate_per_single;
    worksheet.getCell('D21').value = data.margin.golfer_margins.total_fit_rate_per_sharing;

    worksheet.getCell('E20').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_single;
    worksheet.getCell('E21').value = data.margin.non_golfer_margins.total_fit_rate_per_nongolfer_sharing;

    worksheet.getCell('F20').value = data.total_tour_margin.margin_40_total;
    worksheet.getCell('F21').value = data.total_tour_margin.margin_35_total;

    await workbook.xlsx.writeFile(outputPath);
    return outputPath;
};
