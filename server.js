const express = require('express');
const bodyParser = require('body-parser');
const { generateExcelWithDynamicItinerary } = require('./golftours_excel.js');

const app = express();
const PORT = process.env.PORT || 10000;

app.use(bodyParser.json());

app.post('/generate-excel', async (req, res) => {
  try {
    const data = req.body;

    if (!data || !data.itinerary) {
      return res.status(400).json({ error: "Missing required field: itinerary" });
    }

    // Get buffer from generator
    const excelBuffer = await generateExcelWithDynamicItinerary(data);

    // Build email template separately (you can call a helper here or include logic)
    const subject = `Golf Tour Proposal - ${data.lead_name}`;
    const body = `Hi ${data.team_member},\n\nYour AI agent has generated the proposal for ${data.lead_name}.`;

    // Set proper headers for Excel
    const group_filename = (data.lead_name || 'quotation').trim().replace(/\s+/g, "_") + ".xlsx";
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${group_filename}"; filename*=UTF-8''${encodeURIComponent(group_filename)}`
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    // Send Excel file as raw binary
    res.send(excelBuffer);

    // Optionally: return email template as a header (if small) or log it
    // Example: console.log('Email template:', { subject, body });

  } catch (error) {
    console.error("Excel generation error:", error);
    res.status(500).json({ error: "Error generating Excel", details: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
