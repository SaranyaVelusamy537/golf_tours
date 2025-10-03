const express = require('express');
const bodyParser = require('body-parser');
const { generateExcel } = require('./golftours_excel.js');

const app = express();
const PORT = process.env.PORT || 10000;

// Middleware to parse JSON
app.use(bodyParser.json());

// POST endpoint to generate Excel
app.post('/generate-excel', async (req, res) => {
  try {
    const data = req.body;

    // ✅ Ensure itinerary exists
    if (!data || !data.itinerary) {
      return res.status(400).json({ error: "Missing required field: itinerary" });
    }

    const buffer = await generateExcel(data);

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="Quotation Sheet.xlsx"'
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    res.send(buffer);
  } catch (error) {
    console.error("Excel generation error:", error);
    res.status(500).json({ error: "Error generating Excel", details: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
