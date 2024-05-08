import React, { useState } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import "jspdf-autotable";
import "./excelEditor.css";

function ExcelEditor() {
  const [headers, setHeaders] = useState([]);
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [filters, setFilters] = useState([]);
  const [columnTypes, setColumnTypes] = useState([]);

  // Handle file upload and read data
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
      const binaryString = event.target.result;
      const workbook = XLSX.read(binaryString, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Extract headers (first row) and data rows
      const headers = excelData[0];
      const dataRows = excelData.slice(1);

      // Set headers, data, and filtered data
      setHeaders(headers);
      setData(dataRows);
      setFilteredData(dataRows);

      // Initialize filters with empty values
      setFilters(Array(headers.length).fill(""));

      // Determine and set column types
      setColumnTypes(determineColumnTypes(dataRows));
    };

    reader.readAsBinaryString(file);
  };

  // Function to determine column types based on data rows
  const determineColumnTypes = (dataRows) => {
    const types = [];
    if (dataRows.length > 0) {
      dataRows[0].forEach((cell, index) => {
        if (typeof cell === "number") {
          types.push("number");
        } else if (isValidDate(cell)) {
          types.push("date");
        } else {
          types.push("string");
        }
      });
    }
    return types;
  };

  // Helper function to check if a value is a valid date
  const isValidDate = (value) => {
    const date = new Date(value);
    return !isNaN(date.getTime());
  };

  // Apply filters to the data
  const applyFilters = () => {
    const filtered = data.filter((row) =>
      row.every((cell, index) => {  
        if (filters[index] === "") {
          return true;
        }
        return String(cell)
          .toLowerCase()
          .includes(filters[index].toLowerCase());
      })
    );
    setFilteredData(filtered);
  };

  // Handle cell change with data validation
  const handleCellChange = (e, rowIndex, columnIndex) => {
    const newValue = e.target.value;
    const expectedType = columnTypes[columnIndex];

    // Validate the new cell value based on expected data type
    if (!isValidValue(newValue, expectedType)) {
      alert(`Invalid input! Expected a value of type ${expectedType}.`);
      return;
    }

    // Update data state with the new value
    const newData = [...data];
    newData[rowIndex][columnIndex] = newValue;
    setData(newData);

    // Apply filters again to update filtered data if needed
    applyFilters();
  };

  // Function to check if the new value is valid for the expected data type
  const isValidValue = (value, type) => {
    if (type === "number") {
      return !isNaN(parseFloat(value)) && isFinite(value);
    } else if (type === "date") {
      return isValidDate(value);
    } else {
      return true;
    }
  };

  // Function to add a new row
  const addNewRow = () => {
    const newRow = Array(headers.length).fill("");
    const newData = [...data, newRow];
    setData(newData);
    setFilteredData(newData);
  };

  // Export data to Excel file
  const exportToExcel = (dataToExport, fileName) => {
    const combinedData = [headers, ...dataToExport];
    const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, fileName);
  };

  // Export data to JSON file
  const exportToJSON = (dataToExport, fileName) => {
    const jsonData = dataToExport.map((row) => {
      const jsonRow = {};
      row.forEach((cell, index) => {
        const columnHeader = headers[index];
        jsonRow[columnHeader] = cell;
      });
      return jsonRow;
    });

    const jsonString = JSON.stringify(jsonData, null, 2);
    const blob = new Blob([jsonString], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
  };

  // Export data to PDF file
  const exportToPDF = (dataToExport, fileName) => {
    // Create a new instance of jsPDF
    const doc = new jsPDF();

    // Convert the data to a format compatible with jsPDF-autotable
    const dataArray = dataToExport.map((row) =>
      row.map((cell) => String(cell))
    );

    // Use the autoTable method to generate the PDF table
    doc.autoTable({
      head: [headers],
      body: dataArray,
    });

    // Save the PDF file
    doc.save(fileName);
  };

  return (
    <div>
      <h1>Excel Editor</h1>
      <input type="file" onChange={handleFileUpload} />

      {/* Render filter inputs above each column */}
      <div className="filters">
        {filters.map((filter, index) => (
          <input
            key={index}
            type="text"
            placeholder={`Filter column ${index + 1}`}
            value={filter}
            onChange={(e) => {
              const newFilters = [...filters];
              newFilters[index] = e.target.value;
              setFilters(newFilters);
            }}
          />
        ))}
        <button onClick={applyFilters}>Apply Filters</button>
      </div>

      {/* Buttons for exporting data */}
      <button onClick={() => exportToExcel(filteredData, "filtered_data.xlsx")}>
        Export Filtered Data
      </button>
      <button onClick={() => exportToExcel(data, "complete_data.xlsx")}>
        Export Complete Data
      </button>

      {/* Button for exporting to JSON */}
      <button onClick={() => exportToJSON(data, "complete_data.json")}>
        Export Complete Data to JSON
      </button>

      {/* New button for exporting to PDF */}
      <button onClick={() => exportToPDF(filteredData, "filtered_data.pdf")}>
        Export Filtered Data to PDF
      </button>
      <button onClick={() => exportToPDF(data, "complete_data.pdf")}>
        Export Complete Data to PDF
      </button>

      {/* Button to add a new row */}
      <button onClick={addNewRow}>Add New Row</button>

      {/* Render table */}
      <table>
        <thead>
          <tr>
            {headers.map((header, index) => (
              <th key={index}>{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {filteredData.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, columnIndex) => (
                <td key={columnIndex}>
                  <input
                    type="text"
                    value={cell}
                    onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
                  />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default ExcelEditor;
