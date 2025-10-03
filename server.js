// server.js
import express from 'express';
import bodyParser from 'body-parser';
import fs from 'fs';
import { generateExcel } from './golftours_excel.js'; // your existing Excel generator

const app = express();

// Use bodyParser to automatically parse JSON bodies
app.use(bodyParser.json({ limit: '10mb' })); // adjust limit as needed

// Endpoint to generate Excel
app.post('/generate-excel', async (req, res) => {
  try {
    let data = req.body;

    // Safeguard: if body is a string, parse it
    if (typeof data === 'string') {
      data = JSON.parse(data);
    }

    // Check if essential properties exist
    if (!data || !data.itinerary) {
      return res.status(400).json({ error: 'Invalid data: itinerary missing' });
    }

    // Call your existing generateExcel function
    const workbook = await generateExcel(data); // should return a Buffer or Stream

    // Send the Excel file as response
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="golf_tour.xlsx"`
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    // If generateExcel returns a Buffer
    res.send(workbook);

    // If generateExcel returns a stream:
    // workbook.pipe(res);

  } catch (err) {
    console.error('Excel generation error:', err);
    res.status(500).json({ error: 'Error generating Excel', details: err.message });
  }
});

// Start server
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
