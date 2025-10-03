const express = require('express');
const app = express();
const port = process.env.PORT || 10000; // Use Render's PORT or fallback to 10000

// Middleware to parse JSON if needed
app.use(express.json());

// Example route to generate Excel
app.get('/generate-excel', async (req, res) => {
    try {
        const generateExcel = require('./golftours_excel');
        await generateExcel(); // Make sure golftours_excel.js exports a function
        res.send('Excel generated successfully');
    } catch (err) {
        console.error(err);
        res.status(500).send('Error generating Excel');
    }
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
