const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcelWithDynamicItinerary(data) {
  // ---------- Excel generation (unchanged) ----------
  const templatePath = path.join(__dirname, 'public/templates/Golf_Tours_Template.xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const sheet = workbook.getWorksheet('Quotation Sheet');

  // Fill basic info
  sheet.getCell('K5').value = data.lead_name + ' Group';
  sheet.getCell('L5').value = data.team_member;
  sheet.getCell('I16').value = data.golfers;
  sheet.getCell('K16').value = data.non_golfers;

  const dayCellMap = [
    { date: 15, day_of_week:16, hotel: 15, golf: 16, transport: 17 }, // Day 1
    { date: 20, day_of_week:21, hotel: 20, golf: 21, transport: 22 }, // Day 2
    { date: 25, day_of_week:26, hotel: 25, golf: 26, transport: 27 }, // Day 3
    { date: 30, day_of_week:31, hotel: 30, golf: 31, transport: 32 }, // Day 4
    { date: 35, day_of_week:36, hotel: 35, golf: 36, transport: 37 }, // Day 5
    { date: 40, day_of_week:41, hotel: 40, golf: 41, transport: 42 }, // Day 6
    { date: 45, day_of_week:46, hotel: 45, golf: 46, transport: 47 }, // Day 7
    { date: 50, day_of_week:50, hotel: 50, golf: 51, transport: 52 }, // Day 8
    { date: 55, day_of_week:56, hotel: 55, golf: 56, transport: 57 }, // Day 9
    { date: 60, day_of_week:61, hotel: 60, golf: 61, transport: 62 }, // Day 10
    { date: 65, day_of_week:66, hotel: 65, golf: 66, transport: 67 }, // Day 11
    { date: 70, day_of_week:71, hotel: 70, golf: 71, transport: 72 }, // Day 12
  ];

  const itineraryDays = Object.keys(data.itinerary);

  itineraryDays.forEach((dayKey, index) => {
    const dayData = data.itinerary[dayKey];
    const map = dayCellMap[index];
    if (!map) return;

    sheet.getCell(`A${map.date}`).value = dayData.date;
    if (dayData.day_of_week) sheet.getCell(`A${map.day_of_week}`).value = dayData.day_of_week;

    if (dayData.hotel_stay?.length) {
      const hotel = dayData.hotel_stay[0];
      sheet.getCell(`B${map.hotel}`).value = hotel.hotel;
      sheet.getCell(`C${map.hotel}`).value = hotel.Hotel_Sharing;
      sheet.getCell(`D${map.hotel}`).value = hotel.Hotel_Single;
    }

    if (dayData.Golf_round?.length) {
      const golf = dayData.Golf_round[0];
      sheet.getCell(`B${map.golf}`).value = golf.course;
      sheet.getCell(`E${map.golf}`).value = golf.Golf;
    }

    if (dayData.transport?.length) {
      const transport = dayData.transport[0];
      sheet.getCell(`B${map.transport}`).value = transport.transport_type;
      sheet.getCell(`F${map.transport}`).value = transport.rate_per_person;
    }
  });

  const excelBuffer = await workbook.xlsx.writeBuffer();

  // ---------- Email template generation ----------
  function toTwoDecimals(value) {
    return parseFloat(value).toFixed(2);
  }
  function formatDate(dateStr) {
    return dateStr;
  }

  const subject = `Golf Tour Proposal - ${data.lead_name}`;
  
  let groupSize = `${data.golfers} Golfers`;
  if (data.non_golfers && data.non_golfers > 0) {
    groupSize += ` + ${data.non_golfers} Non Golfer${data.non_golfers > 1 ? 's' : ''}`;
  }

  const hotelNights = {};
  Object.values(data.itinerary).forEach(day => {
    const hotelName = day.hotel_stay && day.hotel_stay.length ? day.hotel_stay[0].hotel : null;
    if (hotelName && hotelName.toLowerCase() !== "no hotel") {
      hotelNights[hotelName] = (hotelNights[hotelName] || 0) + 1;
    }
  });
  const accommodationArr = [];
  for (const [hotel, nights] of Object.entries(hotelNights)) {
    accommodationArr.push(`${hotel} for ${nights} Night${nights > 1 ? 's' : ''}`);
  }
  const accommodationLine = `Accommodation: ${accommodationArr.join(' + ')}, all on Bed & Breakfast basis`;

  let golfRoundCount = 0;
  let previousHotel = null;
  const golfRounds = Object.values(data.itinerary)
    .map((day, idx, arr) => {
      const hotelName = day.hotel_stay && day.hotel_stay.length ? day.hotel_stay[0].hotel : null;
      const courseObj = day.Golf_round && day.Golf_round.length ? day.Golf_round[0] : { course: "No Golf", Golf: 0, golf_hint: "Not Available" };
      const courseName = courseObj.course;
      const golfRate = courseObj.Golf || 0;
      const golfHint = courseObj.golf_hint || "";

      let line = `   ⛳ Day #${idx + 1}: `;

      if (idx === 0 && courseName.toLowerCase().includes("no golf")) {
        const airports = data.arrival_airports?.length ? data.arrival_airports.join(" or ") : "Arrival Airport";
        line += `Arrival into ${airports}`;
        if (hotelName && hotelName.toLowerCase() !== "no hotel") {
          line += ` |  Check in at ${hotelName}`;
          previousHotel = hotelName;
        }
      } else if (idx === arr.length - 1) {
        const airports = data.departure_airports?.length ? data.departure_airports.join(" or ") : "Departure Airport";
        line += `Depart from ${airports}`;
        if (golfRate > 0 && !courseName.toLowerCase().includes("no golf")) {
          line += ` |  1 round at ${courseName}`;
          golfRoundCount++;
        } else {
          line += ` |  ${courseName} - Not Available`;
        }
      } else if (golfRate === 0 || golfHint.toLowerCase() === "not available") {
        line += `${courseName} - Not Available`;
        if (hotelName && hotelName.toLowerCase() !== "no hotel" && hotelName !== previousHotel) {
          line += ` |  Check in at ${hotelName}`;
          previousHotel = hotelName;
        }
      } else {
        line += `1 round at ${courseName}`;
        golfRoundCount++;
        if (hotelName && hotelName.toLowerCase() !== "no hotel" && hotelName !== previousHotel) {
          line += ` |  Check in at ${hotelName}`;
          previousHotel = hotelName;
        }
      }

      return line;
    })
    .join('\n');

  const firstTransport = Object.values(data.itinerary)[0].transport[0].transport_type;
  const transfersLine = `🚌 Transfers: ${firstTransport} for duration of trip`;

  let packagePriceLine = `Package Price: €${toTwoDecimals(data.trip_total_margin.golfer_margin_sharing)} per person sharing / €${toTwoDecimals(data.trip_total_margin.golfer_margin_single)} Single`;
  if (data.non_golfers && data.non_golfers > 0) {
    packagePriceLine += `\nPackage Price for Non-Golfers: €${toTwoDecimals(data.trip_total_margin.non_golfer_margin_sharing)} per person sharing / €${toTwoDecimals(data.trip_total_margin.non_golfer_margin_single)} Single`;
  }

  const body = `
Hi ${data.team_member},

Your AI agent is delighted to provide you with the below proposal based on the information provided. Please follow the below steps.

Step 1: Be sure to send this proposal to the client via Pipedrive using the preloaded template, called [GOLF TOURS - NEW QUOTE TEMPLATE].
Step 2: Once the proposal is sent, file this email away in the folder called 'AI Proposals' [2026 Clients (Inbound) – AI Proposals].

Golf Tour Proposal - Ireland
📅 Dates: ${formatDate(data.start_date)} to ${formatDate(data.end_date)}
👥 Group Size: ${groupSize}
🛏️ ${accommodationLine}
🏌️‍♂️ Golf: ${golfRoundCount} rounds
${golfRounds}

${transfersLine}

${packagePriceLine}

Please note the margin charged on this Golf Break is ${data.margin}.

Regards,
Your AI Agent
`;

return {
json: { subject, body },
binary: {
  excelFile: {
    data: excelBuffer,  // raw bytes
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName: `${data.lead_name}_Quotation.xlsx`
  }
  }
};

}

module.exports = { generateExcelWithDynamicItinerary };
