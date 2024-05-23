import React, { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import "jspdf-autotable";
import "./excelEditor.css";
import FileUploader from "./FileUploader";
import Filters from "./Filters";
import SheetSelector from "./SheetSelector";
import AddNewRow from "./AddNewRow";
import ExportButtons from "./ExportButtons";
import DataTable from "./DataTable";

function ExcelEditor() {
  const [workbook, setWorkbook] = useState(null);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState(0);
  const [sheetData, setSheetData] = useState([]);
  const [originalData, setOriginalData] = useState([]);
  const [filters, setFilters] = useState([]);
  const [columnTypes, setColumnTypes] = useState([]);
  const [globalSearch, setGlobalSearch] = useState("");
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);

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

  const handleSheetChange = (index) => {
    setSelectedSheetIndex(index);
    setFilters(Array(sheetData[index]?.data[0]?.length || 0).fill(""));
    setColumnTypes(determineColumnTypes(sheetData[index]?.data || []));
    setActiveSheetIndex(index); // Update active sheet index
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

  return (
    <div className="container">
      <h1 className="main_heading">Excel Editor</h1>

      <FileUploader setWorkbook={setWorkbook} workbook={workbook} />

      <SheetSelector
        sheetData={sheetData}
        handleSheetChange={handleSheetChange}
        activeSheetIndex={activeSheetIndex}
      />

      {workbook && (
        <>
          <Filters
            filters={filters}
            globalSearch={globalSearch}
            selectedSheetIndex={selectedSheetIndex}
            originalData={originalData}
            setSheetData={setSheetData}
            setFilters={setFilters}
            sheetData={sheetData}
            setGlobalSearch={setGlobalSearch}
          />

          <DataTable
            sheetData={sheetData}
            selectedSheetIndex={selectedSheetIndex}
            handleCellChange={handleCellChange}
          />
          <div className="buttons">
            <AddNewRow
              sheetData={sheetData}
              selectedSheetIndex={selectedSheetIndex}
              setSheetData={setSheetData}
              setOriginalData={setOriginalData}
            />

            <br />
            <ExportButtons
              sheetData={sheetData}
              originalData={originalData}
              selectedSheetIndex={selectedSheetIndex}
              workbook={workbook}
            />
          </div>
        </>
      )}
    </div>
  );
}

export default ExcelEditor;
