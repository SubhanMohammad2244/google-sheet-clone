const express = require("express");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json());

let spreadsheetData = Array(5).fill(null).map(() => Array(5).fill("")); // Temporary storage

// API to get spreadsheet data
app.get("/api/data", (req, res) => {
  res.json(spreadsheetData);
});

// API to save spreadsheet data
app.post("/api/data", (req, res) => {
  spreadsheetData = req.body.data;
  res.json({ message: "Data saved successfully!" });
});

app.listen(5000, () => console.log("Server running on port 5000"));
