const express = require('express');
const app = express();
const port = process.env.PORT || 10000;
const generateExcel = require('./golftours_excel');

// Middleware
app.use(express.json());

// POST route for Excel generation
app.post('/generate-excel', async (req, res) => {
    try {
        const data = req.body; // this is your JSON array from n8n
        const filePath = await generateExcel(data); // modify golftours_excel.js to accept data
        res.download(filePath, 'Golf_Tours_Quote.xlsx'); // send as downloadable file
    } catch (err) {
        console.error('Excel generation error:', err);
        res.status(500).send('Error generating Excel');
    }
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
