
import React, { useState, useEffect } from "react";
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

  useEffect(() => {
    if (workbook) {
      const firstSheetName = workbook.SheetNames[0];
      const firstWorksheet = workbook.Sheets[firstSheetName];
      const firstSheetHeaders = XLSX.utils.sheet_to_json(firstWorksheet, { header: 1 })[0];

      const sheetData = workbook.SheetNames.map((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        return { name: sheetName, data: excelData };
      });

      sheetData.forEach((sheet) => {
        if (!sheet.data.some((row) => JSON.stringify(row) === JSON.stringify(firstSheetHeaders))) {
          sheet.data.unshift(firstSheetHeaders);
        }
      });

      setSheetData(sheetData);
      setOriginalData(sheetData.map(sheet => ({ ...sheet, data: [...sheet.data] })));
      setSelectedSheetIndex(0);
      setFilters(Array(sheetData[0]?.data[0]?.length || 0).fill(""));
      setColumnTypes(determineColumnTypes(sheetData[0]?.data || []));
    }
  }, [workbook]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (file instanceof Blob) {
      const reader = new FileReader();

      reader.onload = (event) => {
        const binaryString = event.target.result;
        const wb = XLSX.read(binaryString, { type: "binary" });
        setWorkbook(wb);
      };

      reader.readAsBinaryString(file);
    } else {
      console.error("Invalid file type.");
    }
  };

  const handleSheetChange = (index) => {
    setSelectedSheetIndex(index);
    setFilters(Array(sheetData[index]?.data[0]?.length || 0).fill(""));
    setColumnTypes(determineColumnTypes(sheetData[index]?.data || []));
  };

  const determineColumnTypes = (dataRows) => {
    const types = [];
    if (dataRows.length > 0) {
      dataRows[0].forEach((_, colIndex) => {
        let columnType = 'string';
        for (let rowIndex = 1; rowIndex < dataRows.length; rowIndex++) {
          const cell = dataRows[rowIndex][colIndex];
          if (typeof cell === 'number') {
            columnType = 'number';
            break;
          } else if (isValidDate(cell)) {
            columnType = 'date';
            break;
          }
        }
        types.push(columnType);
      });
    }
    return types;
  };

  const isValidDate = (value) => {
    const date = new Date(value);
    return !isNaN(date.getTime());
  };

  const isValidValue = (value, type) => {
    if (type === 'number') {
      return !isNaN(parseFloat(value)) && isFinite(value);
    } else if (type === 'date') {
      return isValidDate(value);
    } else {
      return true;
    }
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
        if (rowIndex === 0) return true;

        const matchesGlobalSearch = globalSearchEmpty || row.some((cell) =>
          String(cell).toLowerCase().includes(globalSearch.toLowerCase())
        );

        const matchesColumnFilters = row.every((cell, colIndex) => {
          if (filters[colIndex] === "") {
            return true;
          }
          return String(cell).toLowerCase().includes(filters[colIndex].toLowerCase());
        });

        return matchesGlobalSearch && matchesColumnFilters;
      });

      return { ...sheet, data: filteredData };
    });

    setSheetData(filteredSheets);
  };

  const clearFilters = () => {
    setFilters(Array(sheetData[selectedSheetIndex]?.data[0]?.length || 0).fill(""));
    setGlobalSearch("");
    setSheetData(originalData.map(sheet => ({ ...sheet, data: [...sheet.data] })));
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
      newOriginalData[selectedSheetIndex].data[rowIndex][columnIndex] = newValue;
      return newOriginalData;
    });
  };

  const addNewRow = () => {
    const newRow = Array(sheetData[selectedSheetIndex].data[0].length).fill('');
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

//   const exportAllToExcel = () => {
//     const wb = XLSX.utils.book_new();

//     originalData.forEach(({ name, data }) => {
//       const headers = [name];
//       const combinedData = [headers, ...data];
//       const ws = XLSX.utils.aoa_to_sheet(combinedData);
//       XLSX.utils.book_append_sheet(wb, ws, name);
//     });

//     XLSX.writeFile(wb, "exported_data.xlsx");
//   };

// const exportAllToExcel = () => {
//     const wb = XLSX.utils.book_new();
  
//     workbook.SheetNames.forEach(sheetName => {
//       const worksheet = workbook.Sheets[sheetName];
//       const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
//       const ws = XLSX.utils.aoa_to_sheet(data);
//       XLSX.utils.book_append_sheet(wb, ws, sheetName);
//     });
  
//     XLSX.writeFile(wb, "exported_data.xlsx");
//   };
  
// const exportAllToExcel = () => {
//     const wb = XLSX.utils.book_new();
  
//     originalData.forEach(sheet => {
//       const data = sheet.data;
//       const ws = XLSX.utils.aoa_to_sheet(data);
//       XLSX.utils.book_append_sheet(wb, ws, sheet.name);
//     });
  
//     XLSX.writeFile(wb, "exported_data.xlsx");
//   };
  
const exportAllToExcel = () => {
    const wb = XLSX.utils.book_new();
  
    originalData.forEach(sheet => {
      const data = sheet.data.map(row => [...row]); // Deep copy to prevent mutation
      const ws = XLSX.utils.aoa_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
  
    XLSX.writeFile(wb, "exported_data.xlsx");
  };
  

  const exportToJSON = () => {
    const exportedData = originalData.map(sheet => ({
      sheetName: sheet.name,
      data: sheet.data.slice(1).map(row => {
        const obj = {};
        sheet.data[0].forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      })
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
    const exportedData = sheetData.map(sheet => ({
      sheetName: sheet.name,
      data: sheet.data.slice(1).map(row => {
        const obj = {};
        sheet.data[0].forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      })
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
      data: currentSheet.data.slice(1).map(row => {
        const obj = {};
        currentSheet.data[0].forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      })
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

//   const exportCurrentSheetToPDF = () => {
//     const currentSheet = sheetData[selectedSheetIndex];
//     const doc = new jsPDF();
//     doc.autoTable({
//       head: [currentSheet.data[0]],
//       body: currentSheet.data.slice(1),
//     });
//     doc.save(`${currentSheet.name}_filtered.pdf`);
//   };

//   const exportAllSheetsToPDF = () => {
//     const doc = new jsPDF();
//     sheetData.forEach((sheet, index) => {
//       if (index > 0) doc.addPage();
//       doc.text(sheet.name, 10, 10);
//       doc.autoTable({
//         head: [sheet.data[0]],
//         body: sheet.data.slice(1),
//       });
//     });
//     doc.save("filtered_sheets.pdf");
//   };

//   const exportFilteredToPDF = () => {
//     const doc = new jsPDF();
//     sheetData.forEach((sheet, index) => {
//       if (index > 0) doc.addPage();
//       doc.text(sheet.name, 10, 10);
//       doc.autoTable({
//         head: [sheet.data[0]],
//         body: sheet.data.slice(1),
//       });
//     });
//     doc.save("filtered_data.pdf");
//   };

//   const exportFilteredAllToPDF = () => {
//     const doc = new jsPDF();
//     originalData.forEach((sheet, index) => {
//       if (index > 0) doc.addPage();
//       doc.text(sheet.name, 10, 10);
//       doc.autoTable({
//         head: [sheet.data[0]],
//         body: sheet.data.slice(1),
//       });
//     });
//     doc.save("all_filtered_data.pdf");
//   };





const exportToPDF = () => {
    const doc = new jsPDF();
    const sheet = sheetData[selectedSheetIndex];
    const unfilteredData = originalData[selectedSheetIndex].data; // Get the original unfiltered data
    const filteredData = sheet.data || [];
  
    doc.text(sheet.name, 10, 10);
    
    // Use unfilteredData instead of filteredData for PDF generation
    doc.autoTable({
      head: [unfilteredData[0]], // Use headers from unfiltered data
      body: unfilteredData.slice(1), // Use unfiltered data
    });
  
    doc.save(`${sheet.name}_filtered.pdf`);
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
  
    sheetData.forEach(sheet => {
      const sheetName = sheet.name;
      const filteredData = sheet.data || [];
  
      doc.text(sheetName, 10, 10);
      doc.autoTable({
        head: [filteredData[0]], // Use headers from filtered data
        body: filteredData.slice(1), // Use filtered data
      });
      
      doc.addPage(); // Add a new page for each sheet
    });
  
    doc.deletePage(doc.internal.getNumberOfPages()); // Delete the last empty page
  
    doc.save("filtered_all_sheets_data.pdf");
  };


  const exportAllSheetsToPDF = () => {
    const doc = new jsPDF();
  
    originalData.forEach(sheet => {
      const sheetName = sheet.name;
      const sheetData = sheet.data || [];
  
      doc.text(sheetName, 10, 10);
      doc.autoTable({
        head: [sheetData[0]], // Use headers from sheet data
        body: sheetData.slice(1), // Use sheet data
      });
      
      doc.addPage(); // Add a new page for each sheet
    });
  
    doc.deletePage(doc.internal.getNumberOfPages()); // Delete the last empty page
  
    doc.save("all_sheets_data.pdf");
  };

  return (
    <div className="container">
      <h2>Excel Editor</h2>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <div>
        {sheetData.length > 0 && (
          <div className="sheet-selector">
            {sheetData.map((sheet, index) => (
              <button key={index} onClick={() => handleSheetChange(index)}>
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
        </div>
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
                {sheetData[selectedSheetIndex].data[0].map((cell, cellIndex) => (
                  <th key={cellIndex}>{cell}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sheetData[selectedSheetIndex].data.slice(1).map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, cellIndex) => (
                    <td key={cellIndex}>
                      <input
                        type="text"
                        value={cell}
                        onChange={(e) => handleCellChange(e, rowIndex + 1, cellIndex)}
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
        <button onClick={exportFilteredToExcel}>Export All Filtered Sheets to Excel</button>
        <button onClick={exportCurrentSheetFilteredToExcel}>Export Current Sheet Filtered to Excel</button>
        <button onClick={exportAllToExcel}>Export All to Excel</button>
        <button onClick={exportToJSON}>Export All to JSON</button>
        <button onClick={exportCurrentSheetFilteredToJSON}>Export Current Sheet Filtered to JSON</button>
        <button onClick={exportFilteredToJSON}>Export All Sheets Filtered to JSON</button>
        {/* <button onClick={exportCurrentSheetToPDF}>Export Current Sheet to PDF</button>
        <button onClick={exportFilteredToPDF}>Export Current Filtered Data to PDF</button>
        <button onClick={exportFilteredAllToPDF}>Export All Filtered Data to PDF</button>
        <button onClick={exportAllSheetsToPDF}>Export All Sheets Data to PDF</button> */}

        <button onClick={exportToPDF}>Export Current Sheet to PDF</button>
      <button onClick={exportFilteredToPDF}>Export Current Filtered Data to PDF</button>
      
      <button onClick={exportFilteredAllToPDF}>Export All Filtered Data to PDF</button>
     
      <button onClick={exportAllSheetsToPDF}>Export All shhets Data to PDF</button>
      </div>
    </div>
  );
}

export default ExcelEditor;
