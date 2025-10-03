const express = require('express');
const generateExcel = require('./golftours_excel');
const path = require('path');

const app = express();
app.use(express.json());

app.post('/generate-excel', async (req, res) => {
  try {
    const finalJson = req.body; // Input JSON with rates and itinerary
    const outputPath = await generateExcel(finalJson);

    res.download(outputPath, 'Golf_Tours_Generated.xlsx');
  } catch (err) {
    console.error(err);
    res.status(500).send('Error generating Excel');
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
