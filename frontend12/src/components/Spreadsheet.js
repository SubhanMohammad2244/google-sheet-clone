import React, { useState, useEffect } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import "./Spreadsheet.css";

const rows = 5;
const cols = 5;

const Spreadsheet = () => {
  const [data, setData] = useState(
    Array(rows).fill(null).map(() => Array(cols).fill(""))
  );
  const [formula, setFormula] = useState("");

  // Load data from backend when component mounts
  useEffect(() => {
    axios.get("http://localhost:5000/api/data")
      .then((response) => setData(response.data))
      .catch((error) => console.error("Error fetching data:", error));
  }, []);

  // Handle cell input change
  const handleChange = (row, col, value) => {
    const newData = [...data];
    newData[row][col] = value;
    setData(newData);
  };

  // Save data to backend
  const saveData = () => {
    axios.post("http://localhost:5000/api/data", { data })
      .then((response) => alert(response.data.message))
      .catch((error) => alert("Failed to save data."));
  };

  // Function to evaluate formulas
  const evaluateFormula = () => {
    if (formula.startsWith("=")) {
      try {
        const result = evalFormula(formula.substring(1)); // Remove '=' and evaluate
        alert(`Formula Result: ${result}`);
      } catch (error) {
        alert("Invalid Formula!");
      }
    }
  };

  // Support for SUM, AVERAGE, MAX, MIN, COUNT
  const evalFormula = (formula) => {
    if (formula.startsWith("SUM(") && formula.endsWith(")")) {
      return sumFunction(formula);
    } else if (formula.startsWith("AVERAGE(") && formula.endsWith(")")) {
      return avgFunction(formula);
    } else if (formula.startsWith("MAX(") && formula.endsWith(")")) {
      return maxFunction(formula);
    } else if (formula.startsWith("MIN(") && formula.endsWith(")")) {
      return minFunction(formula);
    } else if (formula.startsWith("COUNT(") && formula.endsWith(")")) {
      return countFunction(formula);
    }
    return "Unsupported Formula";
  };

  // Helper functions for formulas
  const sumFunction = (formula) => {
    let range = formula.substring(4, formula.length - 1).split(":");
    return range.reduce((sum, cell) => {
      let [row, col] = parseCell(cell);
      return sum + (parseFloat(data[row][col]) || 0);
    }, 0);
  };

  const avgFunction = (formula) => sumFunction(formula) / countFunction(formula);
  const maxFunction = (formula) => Math.max(...parseRange(formula));
  const minFunction = (formula) => Math.min(...parseRange(formula));
  const countFunction = (formula) => parseRange(formula).length;

  // Convert "A1" to row, col indexes
  const parseCell = (cell) => {
    let col = cell.charCodeAt(0) - 65; // 'A' = 0, 'B' = 1, etc.
    let row = parseInt(cell.substring(1)) - 1;
    return [row, col];
  };

  const parseRange = (formula) => {
    let range = formula.substring(formula.indexOf("(") + 1, formula.indexOf(")")).split(":");
    return range.map(cell => {
      let [row, col] = parseCell(cell);
      return parseFloat(data[row][col]) || 0;
    });
  };

  // Add/Delete Rows and Columns
  const addRow = () => {
    setData([...data, Array(data[0].length).fill("")]);
  };

  const addColumn = () => {
    setData(data.map(row => [...row, ""]));
  };

  // Export and Import Excel Functions
  const exportToExcel = () => {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "spreadsheet.xlsx");
  };

  // FIXED: Corrected file upload handling
  const importFromExcel = (event) => {
    const file = event.target.files[0]; // Get the selected file
    if (!file) {
      alert("No file selected!");
      return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      setData(json);
    };

    reader.readAsArrayBuffer(file); // Ensure correct file reading
  };

  return (
    <div className="spreadsheet-container">
      {/* Formula Bar */}
      <div className="formula-bar">
        <input
          type="text"
          placeholder="Enter formula (e.g., =SUM(A1:A3))"
          value={formula}
          onChange={(e) => setFormula(e.target.value)}
        />
        <button onClick={evaluateFormula}>Apply</button>
      </div>

      {/* Save Data Button */}
      <button onClick={saveData} className="save-button">Save Data</button>

      {/* Add Row & Column Buttons */}
      <button onClick={addRow}>➕ Add Row</button>
      <button onClick={addColumn}>➕ Add Column</button>

      {/* Export & Import Buttons */}
      <button onClick={exportToExcel}>📤 Export to Excel</button>
      <input type="file" accept=".xlsx, .xls" onChange={importFromExcel} />

      {/* Spreadsheet Table */}
      <table className="spreadsheet">
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, colIndex) => (
                <td key={colIndex}>
                  <input
                    type="text"
                    value={cell}
                    onChange={(e) => handleChange(rowIndex, colIndex, e.target.value)}
                  />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Spreadsheet;
