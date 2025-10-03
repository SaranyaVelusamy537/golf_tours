const express = require('express');
const app = express();

const port = process.env.PORT || 10000;

app.use(express.json());

app.get('/generate-excel', async (req, res) => {
  try {
    const generateExcel = require('./golftours_excel');
    const outputPath = await generateExcel();
    res.send(`Excel generated successfully: ${outputPath}`);
  } catch (err) {
    console.error(err);
    res.status(500).send('Error generating Excel');
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
