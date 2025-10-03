app.post('/generate-excel', async (req, res) => {
  try {
    const data = req.body;

    if (!data || !data.itinerary) {
      return res.status(400).json({ error: "Missing required field: itinerary" });
    }

    // ✅ Call the correct function
    const buffer = await generateExcelWithDynamicItinerary(data);

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
