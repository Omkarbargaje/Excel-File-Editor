import React, { useState, useEffect,useCallback } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import "jspdf-autotable";
import "./excelEditor.css";

function ExcelEditor() {
  const [workbook, setWorkbook] = useState(null);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState(0);
  const [sheetData, setSheetData] = useState([]);
  const [originalData, setOriginalData] = useState([]);
  const [filters, setFilters] = useState([]);
  const [columnTypes, setColumnTypes] = useState([]);
  const [globalSearch, setGlobalSearch] = useState("");
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);

 // Define determineColumnTypes function
 const determineColumnTypes = useCallback((dataRows) => {
  const types = [];
  if (dataRows.length > 0) {
    dataRows[0].forEach((_, colIndex) => {
      let columnType = "string";
      for (let rowIndex = 1; rowIndex < dataRows.length; rowIndex++) {
        const cell = dataRows[rowIndex][colIndex];
        if (typeof cell === "number") {
          columnType = "number";
          break;
        } else if (isValidDate(cell)) {
          columnType = "date";
          break;
        }
      }
      types.push(columnType);
    });
  }
  return types;
}, []); // Empty dependency array indicates that this function doesn't depend on any props or state

// useEffect hook
useEffect(() => {
  if (workbook) {
    const sheetData = workbook.SheetNames.map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      return { name: sheetName, data: excelData };
    });

    setSheetData(sheetData);
    setOriginalData([...sheetData]);
    setSelectedSheetIndex(0);
    setFilters(Array(sheetData[0]?.data[0]?.length || 0).fill(""));
    setColumnTypes(determineColumnTypes(sheetData[0]?.data || []));
  }
}, [workbook, determineColumnTypes]); // Include determineColumnTypes in the dependency array




  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const binaryString = event.target.result;
      const wb = XLSX.read(binaryString, { type: "binary" });
      setWorkbook(wb);
    };
    reader.readAsBinaryString(file);
  };


  const isValidDate = (value) => {
    const date = new Date(value);
    return !isNaN(date.getTime());
  };

  const isValidValue = (value, type) => {
    if (type === "number") {
      return !isNaN(parseFloat(value)) && isFinite(value);
    } else if (type === "date") {
      return isValidDate(value);
    } else {
      return true;
    }
  };

  const applyGlobalFilters = () => {
    const allFiltersEmpty = filters.every((filter) => filter === "");
    const globalSearchEmpty = globalSearch.trim() === "";

    if (allFiltersEmpty && globalSearchEmpty) {
      handleSheetChange(selectedSheetIndex);
      return;
    }

    const filteredSheets = originalData.map((sheet) => {
      const filteredData = sheet.data.filter((row, rowIndex) => {
        if (rowIndex === 0) return true; // Keep the header row

        const matchesGlobalSearch =
          globalSearchEmpty ||
          row.some((cell) => {
            if (cell === undefined || cell === null) return false;
            return String(cell)
              .toLowerCase()
              .includes(globalSearch.toLowerCase());
          });

        const matchesRowFilters = row.every((cell) => {
          if (cell === undefined || cell === null) return false;
          return filters.some((filter) =>
            String(cell).toLowerCase().includes(filter.toLowerCase())
          );
        });

        return matchesGlobalSearch && matchesRowFilters;
      });

      return { ...sheet, data: filteredData };
    });

    setSheetData(filteredSheets);
  };

  const applyFilters = () => {
    const allFiltersEmpty = filters.every((filter) => filter === "");
    const globalSearchEmpty = globalSearch.trim() === "";

    if (allFiltersEmpty && globalSearchEmpty) {
      handleSheetChange(selectedSheetIndex);
      return;
    }

    const filteredSheets = originalData.map((sheet) => {
      const filteredData = sheet.data.filter((row, rowIndex) => {
        if (rowIndex === 0) return true; // Keep the header row

        const matchesGlobalSearch =
          globalSearchEmpty ||
          row.some((cell) => {
            if (cell === undefined || cell === null) return false;
            return String(cell)
              .toLowerCase()
              .includes(globalSearch.toLowerCase());
          });

        const matchesRowFilters = row.every((cell, colIndex) => {
          if (cell === undefined || cell === null) return false;
          const filter = filters[colIndex];
          if (filter === "") return true; // No filter applied for this column
          return String(cell).toLowerCase().includes(filter.toLowerCase());
        });

        return matchesGlobalSearch && matchesRowFilters;
      });

      return { ...sheet, data: filteredData };
    });

    setSheetData(filteredSheets);
  };

  const clearFilters = () => {
    setFilters(
      Array(sheetData[selectedSheetIndex]?.data[0]?.length || 0).fill("")
    );
    setGlobalSearch("");
    setSheetData(
      originalData.map((sheet) => ({ ...sheet, data: [...sheet.data] }))
    );
  };

  const handleCellChange = (e, rowIndex, columnIndex) => {
    const newValue = e.target.value;
    const expectedType = columnTypes[columnIndex];

    if (!isValidValue(newValue, expectedType)) {
      alert(`Invalid input! Expected a value of type ${expectedType}.`);
      return;
    }

    setSheetData((prevSheetData) => {
      const newSheetData = [...prevSheetData];
      newSheetData[selectedSheetIndex].data[rowIndex][columnIndex] = newValue;
      return newSheetData;
    });

    setOriginalData((prevOriginalData) => {
      const newOriginalData = [...prevOriginalData];
      newOriginalData[selectedSheetIndex].data[rowIndex][columnIndex] =
        newValue;
      return newOriginalData;
    });
  };

  const addNewRow = () => {
    const newRow = Array(sheetData[selectedSheetIndex].data[0].length).fill("");
    const newData = [...sheetData[selectedSheetIndex].data, newRow];

    setSheetData((prevSheetData) => {
      const newSheetData = [...prevSheetData];
      newSheetData[selectedSheetIndex].data = newData;
      return newSheetData;
    });

    setOriginalData((prevOriginalData) => {
      const newOriginalData = [...prevOriginalData];
      newOriginalData[selectedSheetIndex].data = newData;
      return newOriginalData;
    });
  };

  const exportFilteredToExcel = () => {
    const wb = XLSX.utils.book_new();
    sheetData.forEach((sheet) => {
      const ws = XLSX.utils.aoa_to_sheet(sheet.data);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    XLSX.writeFile(wb, "filtered_data.xlsx");
  };

  const exportAllToExcel = () => {
    const wb = XLSX.utils.book_new();

    originalData.forEach((sheet) => {
      const data = sheet.data.map((row) => [...row]);
      const ws = XLSX.utils.aoa_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });

    XLSX.writeFile(wb, "exported_data.xlsx");
  };

  const exportToJSON = () => {
    const exportedData = originalData.map((sheet) => ({
      sheetName: sheet.name,
      data: sheet.data.slice(1).map((row) => {
        const obj = {};
        sheet.data[0].forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      }),
    }));
    const json = JSON.stringify(exportedData, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "exported_data.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportFilteredToJSON = () => {
    const exportedData = sheetData.map((sheet) => ({
      sheetName: sheet.name,
      data: sheet.data.slice(1).map((row) => {
        const obj = {};
        sheet.data[0].forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      }),
    }));
    const json = JSON.stringify(exportedData, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "filtered_data.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportCurrentSheetFilteredToExcel = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData[selectedSheetIndex].data);
    XLSX.utils.book_append_sheet(wb, ws, sheetData[selectedSheetIndex].name);
    XLSX.writeFile(wb, `${sheetData[selectedSheetIndex].name}_filtered.xlsx`);
  };

  const exportCurrentSheetFilteredToJSON = () => {
    const currentSheet = sheetData[selectedSheetIndex];
    const exportedData = {
      sheetName: currentSheet.name,
      data: currentSheet.data.slice(1).map((row) => {
        const obj = {};
        currentSheet.data[0].forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      }),
    };
    const json = JSON.stringify(exportedData, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${currentSheet.name}_filtered.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    const sheet = sheetData[selectedSheetIndex];
    const unfilteredData = originalData[selectedSheetIndex].data;

    doc.text(sheet.name, 10, 10);

    doc.autoTable({
      head: [unfilteredData[0]],
      body: unfilteredData.slice(1),
    });

    doc.save(`${sheet.name}.pdf`);
  };

  const exportFilteredToPDF = () => {
    const doc = new jsPDF();
    const filteredData = sheetData[selectedSheetIndex]?.data || [];
    doc.autoTable({
      head: [filteredData[0]],
      body: filteredData.slice(1),
    });
    doc.save("filtered_data.pdf");
  };

  const exportFilteredAllToPDF = () => {
    const doc = new jsPDF();

    sheetData.forEach((sheet) => {
      const sheetName = sheet.name;
      const filteredData = sheet.data || [];

      doc.text(sheetName, 10, 10);
      doc.autoTable({
        head: [filteredData[0]],
        body: filteredData.slice(1),
      });

      doc.addPage();
    });

    doc.deletePage(doc.internal.getNumberOfPages());

    doc.save("filtered_all_sheets_data.pdf");
  };

  const exportAllSheetsToPDF = () => {
    const doc = new jsPDF();

    originalData.forEach((sheet) => {
      const sheetName = sheet.name;
      const sheetData = sheet.data || [];

      doc.text(sheetName, 10, 10);
      doc.autoTable({
        head: [sheetData[0]],
        body: sheetData.slice(1),
      });

      doc.addPage();
    });

    doc.deletePage(doc.internal.getNumberOfPages());

    doc.save("all_sheets_data.pdf");
  };

  const downloadOriginalExcel = () => {
    if (!workbook) return;

    const wbBlob = new Blob(
      [s2ab(XLSX.write(workbook, { bookType: "xlsx", type: "binary" }))],
      {
        type: "application/octet-stream",
      }
    );

    const url = URL.createObjectURL(wbBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "original_data.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };

  const handleSheetChange = (index) => {
    setSelectedSheetIndex(index);
    setFilters(Array(sheetData[index]?.data[0]?.length || 0).fill(""));
    setColumnTypes(determineColumnTypes(sheetData[index]?.data || []));
    setActiveSheetIndex(index); // Update active sheet index
  };

  return (
    <div className="container">
      <h1>Excel Editor</h1>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <div>
        {sheetData.length > 0 && (
          <div className="sheet-selector">
            {sheetData.map((sheet, index) => (
              <button
                key={index}
                onClick={() => handleSheetChange(index)}
                className={activeSheetIndex === index ? "active" : ""}
              >
                {sheet.name}
              </button>
            ))}
          </div>
        )}
      </div>
      <div className="filters">
        <div className="global-search">
          <input
            type="text"
            value={globalSearch}
            onChange={(e) => setGlobalSearch(e.target.value)}
            placeholder="Global Search"
          />
          <button onClick={applyGlobalFilters}>Apply Global Filters</button>
        </div>

        <br />
        {filters.map((filter, index) => (
          <input
            key={index}
            type="text"
            value={filter}
            onChange={(e) => {
              const newFilters = [...filters];
              newFilters[index] = e.target.value;
              setFilters(newFilters);
            }}
            placeholder={`Filter column ${index + 1}`}
          />
        ))}
        <button onClick={applyFilters}>Apply Filters</button>
        <button onClick={clearFilters}>Clear Filters</button>
      </div>
      <div className="table-container">
        {sheetData[selectedSheetIndex]?.data.length > 0 && (
          <table>
            <thead>
              <tr>
                {sheetData[selectedSheetIndex].data[0].map(
                  (cell, cellIndex) => (
                    <th key={cellIndex}>{cell}</th>
                  )
                )}
              </tr>
            </thead>
            <tbody>
              {sheetData[selectedSheetIndex].data
                .slice(1)
                .map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.map((cell, cellIndex) => (
                      <td key={cellIndex}>
                        <input
                          type="text"
                          value={cell}
                          onChange={(e) =>
                            handleCellChange(e, rowIndex + 1, cellIndex)
                          }
                        />
                      </td>
                    ))}
                  </tr>
                ))}
            </tbody>
          </table>
        )}
      </div>
      <div className="buttons">
        <button onClick={addNewRow}>Add New Row</button>
        <br />
        <button onClick={exportCurrentSheetFilteredToExcel}>
          Export Current Filtered Sheet to Excel
        </button>
        <button onClick={exportFilteredToExcel}>
          Export All Filtered Sheets to Excel
        </button>
        <button onClick={exportAllToExcel}>Export All to Excel</button>
        <br />
        <button onClick={exportCurrentSheetFilteredToJSON}>
          Export Current Filtered Sheet to JSON
        </button>
        <button onClick={exportFilteredToJSON}>
          Export All Filtered Sheets to JSON
        </button>
        <button onClick={exportToJSON}>Export All Sheet Data to JSON</button>
        <br />
        <button onClick={exportFilteredToPDF}>
          Export Current Sheet Filtered Data to PDF
        </button>
        <button onClick={exportToPDF}>
          Export Current All Sheet Data to PDF
        </button>
        <button onClick={exportFilteredAllToPDF}>
          Export All Filtered Sheets Data to PDF
        </button>
        <button onClick={exportAllSheetsToPDF}>
          Export All Sheets Data to PDF
        </button>
        <br />
        <button onClick={downloadOriginalExcel}>Download Original Excel</button>
      </div>
    </div>
  );
}

export default ExcelEditor;