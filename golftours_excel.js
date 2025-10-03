// Fill itinerary
const dayCellMap = [
  { date: 15, hotel: 15, golf: 16, transport: 17 }, // Day 1
  { date: 20, hotel: 20, golf: 21, transport: 22 }, // Day 2
  { date: 38, hotel: 38, golf: 39, transport: 40 }, // Day 3
  { date: 56, hotel: 56, golf: 57, transport: 58 }, // Day 4
  { date: 74, hotel: 74, golf: 75, transport: 76 }, // Day 5
  // add more if template has more days
];

const itineraryDays = Object.keys(data.itinerary);

itineraryDays.forEach((dayKey, index) => {
  const dayData = data.itinerary[dayKey];
  const map = dayCellMap[index];
  if (!map) return; // skip if no mapping

  // Row 1: Date + Hotel
  sheet.getCell(`A${map.date}`).value = dayData.date; // Date
  if (dayData.hotel_stay && dayData.hotel_stay.length > 0) {
    const hotel = dayData.hotel_stay[0];
    sheet.getCell(`B${map.hotel}`).value = hotel.hotel;           // Hotel Name
    sheet.getCell(`C${map.hotel}`).value = hotel.Hotel_Sharing;   // Hotel Sharing
    sheet.getCell(`D${map.hotel}`).value = hotel.Hotel_Single;    // Hotel Single
  }

  // Row 2: Day of Week + Golf
  if (dayData.Golf_round && dayData.Golf_round.length > 0) {
    const golf = dayData.Golf_round[0];
    sheet.getCell(`B${map.golf}`).value = golf.course; // Golf Club Name
    sheet.getCell(`E${map.golf}`).value = golf.Golf;   // Golf rate/value
  }

  // Row 3: Transport
  if (dayData.transport && dayData.transport.length > 0) {
    const transport = dayData.transport[0];
    sheet.getCell(`B${map.transport}`).value = transport.transport_type; // Transport type
    sheet.getCell(`F${map.transport}`).value = transport.rate_per_person; // Transport rate
  }
});
