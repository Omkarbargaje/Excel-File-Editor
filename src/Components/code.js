import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './excelEditor.css';

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
            const workbook = XLSX.read(binaryString, { type: 'binary' });
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
            setFilters(Array(headers.length).fill(''));

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
                if (typeof cell === 'number') {
                    types.push('number');
                } else if (isValidDate(cell)) {
                    types.push('date');
                } else {
                    types.push('string');
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
        // Filter data based on filters state
        const filtered = data.filter(row =>
            row.every((cell, index) => {
                // If filter is empty, consider it as a match
                if (filters[index] === '') {
                    return true;
                }
                // Perform case-insensitive comparison for the filter
                return String(cell).toLowerCase().includes(filters[index].toLowerCase());
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
        if (type === 'number') {
            return !isNaN(parseFloat(value)) && isFinite(value);
        } else if (type === 'date') {
            return isValidDate(value);
        } else {
            // Default to string type
            return true;
        }
    };

    // Function to add a new row
    const addNewRow = () => {
        // Create a new row with empty values
        const newRow = Array(headers.length).fill('');
        const newData = [...data, newRow];
        setData(newData);
        setFilteredData(newData);
    };

    // Export data to Excel file
    const exportToExcel = (dataToExport, fileName) => {
        const combinedData = [headers, ...dataToExport];
        const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, fileName);
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
                {/* Button to apply filters */}
                <button onClick={applyFilters}>Apply Filters</button>
            </div>

            {/* Buttons to export data */}
            <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
                Export Filtered Data
            </button>
            <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
                Export Complete Data
            </button>

            {/* Button to add a new row */}
            <button onClick={addNewRow}>Add New Row</button>

            {/* Render table */}
            <table>
                <thead>
                    <tr>
                        {/* Render column headers */}
                        {headers.map((header, index) => (
                            <th key={index}>{header}</th>
                        ))}
                    </tr>
                </thead>
                <tbody>
                    {/* Render filtered data */}
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









// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Determine column types and convert date columns
//             const { types, convertedDataRows } = processExcelData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedDataRows);
//             setFilteredData(convertedDataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Set column types
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Process Excel data rows: determine column types and convert date columns
//     const processExcelData = (dataRows) => {
//         const types = [];
//         const convertedDataRows = dataRows.map(row => row.map((cell, index) => {
//             // Determine column types
//             if (Number.isInteger(cell) && cell > 0) {
//                 if (isExcelDate(cell)) {
//                     types[index] = 'date';
//                     return convertExcelDate(cell);
//                 }
//                 types[index] = 'number';
//                 return cell;
//             } else {
//                 types[index] = 'string';
//                 return cell;
//             }
//         }));
//         return { types, convertedDataRows };
//     };

//     // Function to check if a value is an Excel date serial number
//     const isExcelDate = (value) => {
//         // Excel date serial numbers are typically positive integers
//         return Number.isInteger(value) && value > 0;
//     };

//     // Convert Excel date serial number to JavaScript Date object and format
//     const convertExcelDate = (excelDate) => {
//         const baseDate = new Date(Date.UTC(1899, 11, 30));
//         const date = new Date(baseDate.getTime() + excelDate * 86400000);
//         // Return date in 'YYYY-MM-DD' format
//         const formattedDate = date.toISOString().slice(0, 10);
//         return formattedDate;
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         // Filter data based on filters state
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 // If filter is empty, consider it as a match
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 // Perform case-insensitive comparison for the filter
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = newValue;
//         setData(newData);

//         // Apply filters again to update filtered data if needed
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return !isNaN(parseFloat(value)) && isFinite(value);
//         } else if (type === 'date') {
//             return isValidDate(value);
//         } else {
//             // Default to string type
//             return true;
//         }
//     };

//     // Helper function to check if a value is a valid date
//     const isValidDate = (value) => {
//         const date = new Date(value);
//         return !isNaN(date.getTime());
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         // Create a new row with empty values
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     <input
//                                         type="text"
//                                         value={cell}
//                                         onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                     />
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;










// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Process Excel data to determine column types and convert date/number columns
//             const { types, convertedDataRows } = processExcelData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedDataRows);
//             setFilteredData(convertedDataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Set column types
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Process Excel data rows: determine column types and convert date/number columns
//     const processExcelData = (dataRows) => {
//         const types = [];
//         const numRowsToCheck = 10; // Number of rows to check to determine column type
//         const dateThreshold = 0.7; // Proportion of rows needed to classify a column as date
//         const numberThreshold = 0.7; // Proportion of rows needed to classify a column as number
    
//         // Analyze the first few rows of each column to determine data types
//         const columnChecks = Array(dataRows[0].length).fill(null).map(() => ({ dateCount: 0, numCount: 0 }));
    
//         // Check each cell in the first few rows
//         for (let i = 0; i < numRowsToCheck && i < dataRows.length; i++) {
//             dataRows[i].forEach((cell, index) => {
//                 if (isExcelDate(cell)) {
//                     columnChecks[index].dateCount++;
//                 } else if (isNumber(cell)) {
//                     columnChecks[index].numCount++;
//                 }
//             });
//         }
    
//         // Determine column types based on thresholds
//         columnChecks.forEach((check, index) => {
//             if (check.dateCount / numRowsToCheck >= dateThreshold) {
//                 types[index] = 'date';
//             } else if (check.numCount / numRowsToCheck >= numberThreshold) {
//                 types[index] = 'number';
//             } else {
//                 types[index] = 'string';
//             }
//         });
    
//         // Convert data rows based on determined column types
//         const convertedDataRows = dataRows.map(row =>
//             row.map((cell, index) => {
//                 if (types[index] === 'date') {
//                     // Convert date cell
//                     return convertExcelDate(cell);
//                 } else if (types[index] === 'number') {
//                     // Convert number cell
//                     return parseFloat(cell);
//                 } else {
//                     // Return string cell as-is
//                     return cell;
//                 }
//             })
//         );
    
//         return { types, convertedDataRows };
//     };

//     // Function to check if a value is an Excel date serial number
//     const isExcelDate = (cell) => {
//         if (typeof cell === 'number' && cell > 0) {
//             const baseDate = new Date(Date.UTC(1899, 11, 30));
//             const date = new Date(baseDate.getTime() + cell * 86400000);
//             // Verify if the calculated date is valid
//             if (!isNaN(date.getTime()) && cell >= 0 && cell <= 2958465) { // Excel date range
//                 return true;
//             }
//         }
//         return false;
//     };

//     // Convert Excel date serial number to JavaScript Date object and format
//     const convertExcelDate = (excelDate) => {
//         const baseDate = new Date(Date.UTC(1899, 11, 30));
//         const date = new Date(baseDate.getTime() + excelDate * 86400000);
//         // Return date in 'YYYY-MM-DD' format
//         return date.toISOString().slice(0, 10);
//     };

//     // Function to check if a value is a valid number
//     const isNumber = (cell) => {
//         // Check if the cell can be parsed as a number or is already a number
//         if (typeof cell === 'number') {
//             return true;
//         } else {
//             const parsedNumber = parseFloat(cell);
//             return !isNaN(parsedNumber) && isFinite(parsedNumber);
//         }
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         // Filter data based on filters state
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 // If filter is empty, consider it as a match
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 // Perform case-insensitive comparison for the filter
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];
    
//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }
    
//         // Convert input value to the expected data type
//         let convertedValue;
    
//         if (expectedType === 'number') {
//             convertedValue = parseFloat(newValue);
//         } else if (expectedType === 'date') {
//             // Validate the input date and convert to 'YYYY-MM-DD' format
//             const date = new Date(newValue);
//             if (isNaN(date.getTime())) {
//                 alert('Invalid date format! Please enter a valid date in the format YYYY-MM-DD.');
//                 return;
//             }
//             convertedValue = date.toISOString().slice(0, 10);
//         } else {
//             convertedValue = newValue;
//         }
    
//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = convertedValue;
//         setData(newData);
    
//         // Apply filters again to update filtered data
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return isNumber(value);
//         } else if (type === 'date') {
//             // Check if value is a valid date
//             return !isNaN(new Date(value).getTime());
//         } else {
//             // Default to string type
//             return true;
//         }
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         // Create a new row with empty values
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     <input
//                                         type={columnTypes[columnIndex] === 'number' ? 'number' : 'text'}
//                                         value={cell}
//                                         onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                     />
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;







// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Process Excel data to determine column types and convert date/number columns
//             const { types, convertedDataRows } = processExcelData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedDataRows);
//             setFilteredData(convertedDataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Set column types
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Process Excel data rows: determine column types and convert date/number columns
//     const processExcelData = (dataRows) => {
//         const types = [];
//         const numRowsToCheck = 10; // Number of rows to check to determine column type
//         const dateThreshold = 0.7; // Proportion of rows needed to classify a column as date
//         const numberThreshold = 0.7; // Proportion of rows needed to classify a column as number
    
//         // Analyze the first few rows of each column to determine data types
//         const columnChecks = Array(dataRows[0].length).fill(null).map(() => ({ dateCount: 0, numCount: 0 }));
    
//         // Check each cell in the first few rows
//         for (let i = 0; i < numRowsToCheck && i < dataRows.length; i++) {
//             dataRows[i].forEach((cell, index) => {
//                 if (isExcelDate(cell)) {
//                     columnChecks[index].dateCount++;
//                 } else if (isNumber(cell)) {
//                     columnChecks[index].numCount++;
//                 }
//             });
//         }
    
//         // Determine column types based on thresholds
//         columnChecks.forEach((check, index) => {
//             if (check.dateCount / numRowsToCheck >= dateThreshold) {
//                 types[index] = 'date';
//             } else if (check.numCount / numRowsToCheck >= numberThreshold) {
//                 types[index] = 'number';
//             } else {
//                 types[index] = 'string';
//             }
//         });
    
//         // Convert data rows based on determined column types
//         const convertedDataRows = dataRows.map(row =>
//             row.map((cell, index) => {
//                 if (types[index] === 'date') {
//                     return convertExcelDate(cell);
//                 } else if (types[index] === 'number') {
//                     return parseFloat(cell);
//                 } else {
//                     return cell;
//                 }
//             })
//         );
    
//         return { types, convertedDataRows };
//     };

//     // Function to check if a value is an Excel date serial number
//     const isExcelDate = (cell) => {
//         if (typeof cell === 'number' && cell > 0) {
//             const baseDate = new Date(Date.UTC(1899, 11, 30));
//             const date = new Date(baseDate.getTime() + cell * 86400000);
//             // Verify if the calculated date is valid
//             if (!isNaN(date.getTime()) && cell >= 0 && cell <= 2958465) { // Excel date range
//                 return true;
//             }
//         }
//         return false;
//     };

//     // Convert Excel date serial number to JavaScript Date object and format
//     const convertExcelDate = (excelDate) => {
//         const baseDate = new Date(Date.UTC(1899, 11, 30));
//         const date = new Date(baseDate.getTime() + excelDate * 86400000);
//         // Return date in 'YYYY-MM-DD' format
//         return date.toISOString().slice(0, 10);
//     };

//     // Function to check if a value is a valid number
//     const isNumber = (cell) => {
//         // Check if the cell can be parsed as a number or is already a number
//         if (typeof cell === 'number') {
//             return true;
//         } else {
//             const parsedNumber = parseFloat(cell);
//             return !isNaN(parsedNumber) && isFinite(parsedNumber);
//         }
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         // Filter data based on filters state
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 // If filter is empty, consider it as a match
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 // Perform case-insensitive comparison for the filter
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // Convert input value to the expected data type
//         let convertedValue;

//         if (expectedType === 'number') {
//             // Convert value to number
//             convertedValue = parseFloat(newValue);
//         } else if (expectedType === 'date') {
//             // Convert and validate the date
//             const date = new Date(newValue);
//             if (isNaN(date.getTime())) {
//                 alert('Invalid date format! Please enter a valid date in the format YYYY-MM-DD.');
//                 return;
//             }
//             convertedValue = date.toISOString().slice(0, 10);
//         } else {
//             // Otherwise, just treat as string
//             convertedValue = newValue;
//         }

//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = convertedValue;
//         setData(newData);

//         // Apply filters again to update filtered data
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return isNumber(value);
//         } else if (type === 'date') {
//             // Check if value is a valid date
//             return !isNaN(new Date(value).getTime());
//         } else {
//             // Default to string type
//             return true;
//         }
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         // Create a new row with empty values
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     {columnTypes[columnIndex] === 'date' ? (
//                                         <input
//                                             type="date" // Use the 'date' input type
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     ) : (
//                                         <input
//                                             type={columnTypes[columnIndex] === 'number' ? 'number' : 'text'}
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     )}
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;








// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Process Excel data to determine column types and convert date/number columns
//             const { types, convertedDataRows } = processExcelData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedDataRows);
//             setFilteredData(convertedDataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Set column types
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Process Excel data rows: determine column types and convert date/number columns
//     const processExcelData = (dataRows) => {
//         const types = [];
//         const numRowsToCheck = 10; // Number of rows to check to determine column type
//         const dateThreshold = 0.7; // Proportion of rows needed to classify a column as date
//         const numberThreshold = 0.7; // Proportion of rows needed to classify a column as number

//         // Analyze the first few rows of each column to determine data types
//         const columnChecks = Array(dataRows[0].length).fill(null).map(() => ({ dateCount: 0, numCount: 0 }));

//         // Check each cell in the first few rows
//         for (let i = 0; i < numRowsToCheck && i < dataRows.length; i++) {
//             dataRows[i].forEach((cell, index) => {
//                 if (isExcelDate(cell)) {
//                     columnChecks[index].dateCount++;
//                 } else if (isNumber(cell)) {
//                     columnChecks[index].numCount++;
//                 }
//             });
//         }

//         // Determine column types based on thresholds
//         columnChecks.forEach((check, index) => {
//             if (check.dateCount / numRowsToCheck >= dateThreshold) {
//                 types[index] = 'date';
//             } else if (check.numCount / numRowsToCheck >= numberThreshold) {
//                 types[index] = 'number';
//             } else {
//                 types[index] = 'string';
//             }
//         });

//         // Convert data rows based on determined column types
//         const convertedDataRows = dataRows.map(row =>
//             row.map((cell, index) => {
//                 if (types[index] === 'date') {
//                     // Convert date cell
//                     return convertExcelDate(cell);
//                 } else if (types[index] === 'number') {
//                     // Convert number cell
//                     return parseFloat(cell);
//                 } else {
//                     // Return string cell as-is
//                     return cell;
//                 }
//             })
//         );

//         return { types, convertedDataRows };
//     };

//     // Function to check if a value is an Excel date serial number
//     const isExcelDate = (cell) => {
//         if (typeof cell === 'number' && cell > 0 && cell <= 2958465) {
//             const date = convertExcelDate(cell);
//             // Check if date is within a reasonable range (e.g., year range 1900-2100)
//             const dateObj = new Date(date);
//             return dateObj.getFullYear() >= 1900 && dateObj.getFullYear() <= 2100;
//         }
//         return false;
//     };

//     // Convert Excel date serial number to JavaScript Date object and format
//     const convertExcelDate = (excelDate) => {
//         const baseDate = new Date(Date.UTC(1899, 11, 30));
//         const date = new Date(baseDate.getTime() + excelDate * 86400000);
//         // Return date in 'YYYY-MM-DD' format
//         return date.toISOString().slice(0, 10);
//     };

//     // Function to check if a value is a valid number
//     const isNumber = (cell) => {
//         // Check if the cell can be parsed as a number or is already a number
//         if (typeof cell === 'number') {
//             return true;
//         } else {
//             const parsedNumber = parseFloat(cell);
//             return !isNaN(parsedNumber) && isFinite(parsedNumber);
//         }
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         // Filter data based on filters state
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 // If filter is empty, consider it as a match
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 // Perform case-insensitive comparison for the filter
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // Convert input value to the expected data type
//         let convertedValue;

//         if (expectedType === 'number') {
//             // Convert value to number
//             convertedValue = parseFloat(newValue);
//         } else if (expectedType === 'date') {
//             // Convert and validate the date
//             const date = new Date(newValue);
//             if (isNaN(date.getTime())) {
//                 alert('Invalid date format! Please enter a valid date in the format YYYY-MM-DD.');
//                 return;
//             }
//             convertedValue = date.toISOString().slice(0, 10);
//         } else {
//             // Otherwise, just treat as string
//             convertedValue = newValue;
//         }

//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = convertedValue;
//         setData(newData);

//         // Apply filters again to update filtered data
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return isNumber(value);
//         } else if (type === 'date') {
//             // Check if value is a valid date
//             return !isNaN(new Date(value).getTime());
//         } else {
//             // Default to string type
//             return true;
//         }
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         // Create a new row with empty values
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     {columnTypes[columnIndex] === 'date' ? (
//                                         <input
//                                             type="date" // Use the 'date' input type
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     ) : (
//                                         <input
//                                             type={columnTypes[columnIndex] === 'number' ? 'number' : 'text'}
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     )}
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;











// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Process Excel data to determine column types and convert date/number columns
//             const { types, convertedDataRows } = processExcelData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedDataRows);
//             setFilteredData(convertedDataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Set column types
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Process Excel data rows: determine column types and convert date/number columns
//     const processExcelData = (dataRows) => {
//         const types = [];
//         const numRowsToCheck = 10; // Number of rows to check to determine column type
//         const dateThreshold = 0.7; // Proportion of rows needed to classify a column as date
//         const numberThreshold = 0.7; // Proportion of rows needed to classify a column as number
    
//         // Analyze the first few rows of each column to determine data types
//         const columnChecks = Array(dataRows[0].length).fill(null).map(() => ({ dateCount: 0, numCount: 0 }));
    
//         // Check each cell in the first few rows
//         for (let i = 0; i < numRowsToCheck && i < dataRows.length; i++) {
//             dataRows[i].forEach((cell, index) => {
//                 if (isExcelDate(cell)) {
//                     columnChecks[index].dateCount++;
//                 } else if (isNumber(cell)) {
//                     columnChecks[index].numCount++;
//                 }
//             });
//         }
    
//         // Determine column types based on thresholds
//         columnChecks.forEach((check, index) => {
//             if (check.dateCount / numRowsToCheck >= dateThreshold) {
//                 types[index] = 'date';
//             } else if (check.numCount / numRowsToCheck >= numberThreshold) {
//                 types[index] = 'number';
//             } else {
//                 types[index] = 'string';
//             }
//         });
    
//         // Convert data rows based on determined column types
//         const convertedDataRows = dataRows.map(row =>
//             row.map((cell, index) => {
//                 if (types[index] === 'date') {
//                     // Convert date cell
//                     return convertExcelDate(cell);
//                 } else if (types[index] === 'number') {
//                     // Convert number cell
//                     return parseFloat(cell);
//                 } else {
//                     // Return string cell as-is
//                     return cell;
//                 }
//             })
//         );
    
//         return { types, convertedDataRows };
//     };

//     // Function to check if a value is an Excel date serial number
//     const isExcelDate = (cell) => {
//         if (typeof cell === 'number' && cell > 0) {
//             const baseDate = new Date(Date.UTC(1899, 11, 30));
//             const date = new Date(baseDate.getTime() + cell * 86400000);
//             // Verify if the calculated date is valid
//             if (!isNaN(date.getTime()) && cell >= 0 && cell <= 2958465) { // Excel date range
//                 return true;
//             }
//         }
//         return false;
//     };

//     // Convert Excel date serial number to JavaScript Date object and format
//     const convertExcelDate = (excelDate) => {
//         const baseDate = new Date(Date.UTC(1899, 11, 30));
//         const date = new Date(baseDate.getTime() + excelDate * 86400000);
//         // Return date in 'YYYY-MM-DD' format
//         return date.toISOString().slice(0, 10);
//     };

//     // Function to check if a value is a valid number
//     const isNumber = (cell) => {
//         // Check if the cell can be parsed as a number or is already a number
//         if (typeof cell === 'number') {
//             return true;
//         } else {
//             const parsedNumber = parseFloat(cell);
//             return !isNaN(parsedNumber) && isFinite(parsedNumber);
//         }
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         // Filter data based on filters state
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 // If filter is empty, consider it as a match
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 // Perform case-insensitive comparison for the filter
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];
    
//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }
    
//         // Convert input value to the expected data type
//         let convertedValue;
    
//         if (expectedType === 'number') {
//             convertedValue = parseFloat(newValue);
//         } else if (expectedType === 'date') {
//             // Validate the input date and convert to 'YYYY-MM-DD' format
//             const date = new Date(newValue);
//             if (isNaN(date.getTime())) {
//                 alert('Invalid date format! Please enter a valid date in the format YYYY-MM-DD.');
//                 return;
//             }
//             convertedValue = date.toISOString().slice(0, 10);
//         } else {
//             convertedValue = newValue;
//         }
    
//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = convertedValue;
//         setData(newData);
    
//         // Apply filters again to update filtered data
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return isNumber(value);
//         } else if (type === 'date') {
//             // Check if value is a valid date
//             return !isNaN(new Date(value).getTime());
//         } else {
//             // Default to string type
//             return true;
//         }
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         // Create a new row with empty values
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     <input
//                                         type={columnTypes[columnIndex] === 'number' ? 'number' : 'text'}
//                                         value={cell}
//                                         onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                     />
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;













// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Determine column types and process data rows
//             const { types, convertedData } = determineColumnTypesAndConvertData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedData);
//             setFilteredData(convertedData);
//             setFilters(Array(headers.length).fill(''));
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Determine column types and convert data
//     const determineColumnTypesAndConvertData = (dataRows) => {
//         const numRowsToCheck = 10; // Number of rows to check to determine column type
//         const columnCount = dataRows[0].length;
//         const columnTypes = Array(columnCount).fill('string'); // Default type is string
//         const columnChecks = Array.from({ length: columnCount }, () => ({
//             dateCount: 0,
//             numCount: 0,
//         }));

//         // Analyze first few rows to determine types
//         for (let i = 0; i < numRowsToCheck && i < dataRows.length; i++) {
//             dataRows[i].forEach((cell, index) => {
//                 if (isDateString(cell)) {
//                     columnChecks[index].dateCount++;
//                 } else if (isNumber(cell)) {
//                     columnChecks[index].numCount++;
//                 }
//             });
//         }

//         // Classify columns based on highest proportion
//         columnChecks.forEach((check, index) => {
//             if (check.dateCount > check.numCount && check.dateCount > numRowsToCheck / 2) {
//                 columnTypes[index] = 'date';
//             } else if (check.numCount > check.dateCount && check.numCount > numRowsToCheck / 2) {
//                 columnTypes[index] = 'number';
//             }
//         });

//         // Convert data rows based on determined column types
//         const convertedData = dataRows.map((row) => row.map((cell, index) => convertCellValue(cell, columnTypes[index])));

//         return { types: columnTypes, convertedData };
//     };

//     // Check if the cell is a valid date string
//     const isDateString = (cell) => {
//         if (typeof cell === 'string') {
//             const date = new Date(cell);
//             return !isNaN(date.getTime());
//         }
//         return false;
//     };

//     // Check if the cell is a valid number
//     const isNumber = (cell) => {
//         if (typeof cell === 'number') {
//             return true;
//         }
//         const parsedNumber = parseFloat(cell);
//         return !isNaN(parsedNumber) && isFinite(parsedNumber);
//     };

//     // Convert cell value based on column type
//     const convertCellValue = (cell, type) => {
//         if (type === 'date' && typeof cell === 'string') {
//             const date = new Date(cell);
//             if (!isNaN(date.getTime())) {
//                 return date.toISOString().slice(0, 10);
//             }
//         } else if (type === 'number') {
//             const parsedNumber = parseFloat(cell);
//             if (!isNaN(parsedNumber)) {
//                 return parsedNumber;
//             }
//         }
//         return cell;
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         const filtered = data.filter((row) =>
//             row.every((cell, index) => {
//                 const filter = filters[index].toLowerCase();
//                 return !filter || String(cell).toLowerCase().includes(filter);
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // Convert input value to the expected data type
//         let convertedValue;

//         if (expectedType === 'number') {
//             convertedValue = parseFloat(newValue);
//         } else if (expectedType === 'date') {
//             const date = new Date(newValue);
//             if (isNaN(date.getTime())) {
//                 alert('Invalid date format! Please enter a valid date in the format YYYY-MM-DD.');
//                 return;
//             }
//             convertedValue = date.toISOString().slice(0, 10);
//         } else {
//             convertedValue = newValue;
//         }

//         // Update data state with the new value
//         const updatedData = [...data];
//         updatedData[rowIndex][columnIndex] = convertedValue;
//         setData(updatedData);
//         applyFilters(); // Apply filters after updating data
//     };

//     // Check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return isNumber(value);
//         } else if (type === 'date') {
//             return isDateString(value);
//         }
//         return true; // String type
//     };

//     // Add a new row with empty cells
//     const addNewRow = () => {
//         const newRow = Array(headers.length).fill('');
//         setData((prevData) => [...prevData, newRow]);
//         setFilteredData((prevData) => [...prevData, newRow]);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Filters */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Export buttons */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>Export Filtered Data</button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>Export Complete Data</button>

//             {/* Add new row button */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Data table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     {columnTypes[columnIndex] === 'date' ? (
//                                         <input
//                                             type="date"
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     ) : columnTypes[columnIndex] === 'number' ? (
//                                         <input
//                                             type="number"
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     ) : (
//                                         <input
//                                             type="text"
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     )}
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;













// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Determine column types and convert data
//             const { types, convertedData } = determineColumnTypesAndConvertData(dataRows);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(convertedData);
//             setFilteredData(convertedData);
//             setFilters(Array(headers.length).fill(''));
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Determine column types and convert data
//     const determineColumnTypesAndConvertData = (dataRows) => {
//         const numRowsToCheck = 10; // Number of rows to check to determine column type
//         const columnCount = dataRows[0].length;
//         const columnTypes = Array(columnCount).fill('string'); // Default type is string
//         const columnChecks = Array.from({ length: columnCount }, () => ({
//             dateCount: 0,
//             numCount: 0,
//         }));

//         // Analyze first few rows to determine types
//         for (let i = 0; i < numRowsToCheck && i < dataRows.length; i++) {
//             dataRows[i].forEach((cell, index) => {
//                 if (isDateString(cell)) {
//                     columnChecks[index].dateCount++;
//                 } else if (isNumber(cell)) {
//                     columnChecks[index].numCount++;
//                 }
//             });
//         }

//         // Classify columns based on highest proportion
//         columnChecks.forEach((check, index) => {
//             if (check.dateCount > check.numCount && check.dateCount > numRowsToCheck / 2) {
//                 columnTypes[index] = 'date';
//             } else if (check.numCount > check.dateCount && check.numCount > numRowsToCheck / 2) {
//                 columnTypes[index] = 'number';
//             }
//         });

//         // Convert data rows based on determined column types
//         const convertedData = dataRows.map(row =>
//             row.map((cell, index) => convertCellValue(cell, columnTypes[index]))
//         );

//         return { types: columnTypes, convertedData };
//     };

//     // Check if the cell is a valid date string
//     const isDateString = (cell) => {
//         if (typeof cell === 'string' || typeof cell === 'number') {
//             const date = new Date(cell);
//             return !isNaN(date.getTime());
//         }
//         return false;
//     };

//     // Check if the cell is a valid number
//     const isNumber = (cell) => {
//         const parsedNumber = parseFloat(cell);
//         return !isNaN(parsedNumber) && isFinite(parsedNumber);
//     };

//     // Convert cell value based on column type
//     const convertCellValue = (cell, type) => {
//         if (type === 'date') {
//             const date = new Date(cell);
//             if (!isNaN(date.getTime())) {
//                 return date.toISOString().slice(0, 10); // Convert to YYYY-MM-DD format
//             }
//         } else if (type === 'number') {
//             const parsedNumber = parseFloat(cell);
//             if (!isNaN(parsedNumber)) {
//                 return parsedNumber;
//             }
//         }
//         return cell; // Otherwise, return as is
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         const filtered = data.filter((row) =>
//             row.every((cell, index) => {
//                 const filter = filters[index].toLowerCase();
//                 return !filter || String(cell).toLowerCase().includes(filter);
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // Convert input value to the expected data type
//         let convertedValue;

//         if (expectedType === 'number') {
//             convertedValue = parseFloat(newValue);
//         } else if (expectedType === 'date') {
//             const date = new Date(newValue);
//             if (isNaN(date.getTime())) {
//                 alert('Invalid date format! Please enter a valid date in the format YYYY-MM-DD.');
//                 return;
//             }
//             convertedValue = date.toISOString().slice(0, 10);
//         } else {
//             convertedValue = newValue;
//         }

//         // Update data state with the new value
//         const updatedData = [...data];
//         updatedData[rowIndex][columnIndex] = convertedValue;
//         setData(updatedData);
//         applyFilters(); // Apply filters after updating data
//     };

//     // Check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return isNumber(value);
//         } else if (type === 'date') {
//             return isDateString(value);
//         }
//         return true; // For string type
//     };

//     // Add a new row with empty cells
//     const addNewRow = () => {
//         const newRow = Array(headers.length).fill('');
//         setData((prevData) => [...prevData, newRow]);
//         setFilteredData((prevData) => [...prevData, newRow]);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Filters */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Export buttons */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>Export Filtered Data</button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>Export Complete Data</button>

//             {/* Add new row button */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Data table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => (
//                                 <td key={columnIndex}>
//                                     {columnTypes[columnIndex] === 'date' ? (
//                                         <input
//                                             type="date"
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     ) : columnTypes[columnIndex] === 'number' ? (
//                                         <input
//                                             type="number"
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     ) : (
//                                         <input
//                                             type="text"
//                                             value={cell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     )}
//                                 </td>
//                             ))}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;



















// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);
//     const [dateColumns, setDateColumns] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(dataRows);
//             setFilteredData(dataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Determine and set column types and date columns
//             setColumnTypes(determineColumnTypes(dataRows));
//             setDateColumns(determineDateColumns(headers, worksheet));
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Function to determine column types based on data rows
//     const determineColumnTypes = (dataRows) => {
//         const types = [];
//         if (dataRows.length > 0) {
//             dataRows[0].forEach((cell, index) => {
//                 if (typeof cell === 'number') {
//                     types.push('number');
//                 } else if (isValidDate(cell)) {
//                     types.push('date');
//                 } else {
//                     types.push('string');
//                 }
//             });
//         }
//         return types;
//     };

//     // Function to determine date columns and their formats
//     const determineDateColumns = (headers, worksheet) => {
//         const dateColumns = [];
//         headers.forEach((header, index) => {
//             const firstCell = worksheet[XLSX.utils.encode_cell({ c: index, r: 1 })]; // Cell in first data row
//             if (firstCell && firstCell.z) {
//                 const format = firstCell.z;
//                 if (format.includes('d') || format.includes('m') || format.includes('y')) {
//                     dateColumns.push({ index, format });
//                 }
//             }
//         });
//         return dateColumns;
//     };

//     // Helper function to check if a value is a valid date
//     const isValidDate = (value) => {
//         const date = new Date(value);
//         return !isNaN(date.getTime());
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         // Filter data based on filters state
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 // If filter is empty, consider it as a match
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 // Perform case-insensitive comparison for the filter
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation and date formatting
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         const newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // If the column is a date column, try parsing and formatting the date
//         if (expectedType === 'date') {
//             const dateColumn = dateColumns.find(col => col.index === columnIndex);
//             if (dateColumn) {
//                 const parsedDate = new Date(newValue);
//                 if (!isNaN(parsedDate)) {
//                     // Format the date according to the expected format
//                     newValue = XLSX.SSF.format(dateColumn.format, parsedDate);
//                 }
//             }
//         }

//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = newValue;
//         setData(newData);

//         // Apply filters again to update filtered data if needed
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return !isNaN(parseFloat(value)) && isFinite(value);
//         } else if (type === 'date') {
//             return isValidDate(value);
//         } else {
//             // Default to string type
//             return true;
//         }
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         // Create a new row with empty values
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);

//         // Apply date formats to worksheet
//         dateColumns.forEach(({ index, format }) => {
//             const colRef = XLSX.utils.encode_col(index);
//             worksheet[`!cols`] = worksheet[`!cols`] || [];
//             worksheet[`!cols`][index] = { z: format };
//         });

//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => {
//                                 const expectedType = columnTypes[columnIndex];
//                                 const isDateColumn = expectedType === 'date';
//                                 let formattedCell = cell;
//                                 if (isDateColumn && cell) {
//                                     const dateColumn = dateColumns.find(col => col.index === columnIndex);
//                                     if (dateColumn) {
//                                         const parsedDate = new Date(cell);
//                                         formattedCell = !isNaN(parsedDate) ? XLSX.SSF.format(dateColumn.format, parsedDate) : cell;
//                                     }
//                                 }

//                                 return (
//                                     <td key={columnIndex}>
//                                         <input
//                                             type="text"
//                                             value={formattedCell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     </td>
//                                 );
//                             })}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;














// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);
//     const [dateColumns, setDateColumns] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(dataRows);
//             setFilteredData(dataRows);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Determine column types and date columns
//             setColumnTypes(determineColumnTypes(dataRows));
//             setDateColumns(determineDateColumns(headers, worksheet));
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Function to determine column types based on data rows
//     const determineColumnTypes = (dataRows) => {
//         const types = [];
//         if (dataRows.length > 0) {
//             dataRows[0].forEach((cell, index) => {
//                 if (typeof cell === 'number') {
//                     types.push('number');
//                 } else if (isValidDate(cell)) {
//                     types.push('date');
//                 } else {
//                     types.push('string');
//                 }
//             });
//         }
//         return types;
//     };

//     // Function to determine date columns and their formats
//     const determineDateColumns = (headers, worksheet) => {
//         const dateColumns = [];
//         headers.forEach((header, index) => {
//             const col = XLSX.utils.encode_col(index);
//             const dateFormat = worksheet[col + '2']; // Check the format in the first data row
//             if (dateFormat && dateFormat.z && (dateFormat.z.includes('d') || dateFormat.z.includes('m') || dateFormat.z.includes('y'))) {
//                 dateColumns.push({ index, format: dateFormat.z });
//             }
//         });
//         return dateColumns;
//     };

//     // Helper function to check if a value is a valid date
//     const isValidDate = (value) => {
//         const date = new Date(value);
//         return !isNaN(date.getTime());
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation and date formatting
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         let newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate the new cell value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // If the column is a date column, try parsing and formatting the date
//         if (expectedType === 'date') {
//             const parsedDate = new Date(newValue);
//             if (!isNaN(parsedDate)) {
//                 const dateColumn = dateColumns.find(col => col.index === columnIndex);
//                 if (dateColumn) {
//                     // Use XLSX SSF format to convert the date to the expected format
//                     newValue = XLSX.SSF.format(dateColumn.format, parsedDate);
//                 }
//             }
//         }

//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = newValue;
//         setData(newData);
//         applyFilters();
//     };

//     // Function to check if the new value is valid for the expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return !isNaN(parseFloat(value)) && isFinite(value);
//         } else if (type === 'date') {
//             return isValidDate(value);
//         } else {
//             return true;
//         }
//     };

//     // Function to add a new row
//     const addNewRow = () => {
//         const newRow = Array(headers.length).fill('');
//         const newData = [...data, newRow];
//         setData(newData);
//         setFilteredData(newData);
//     };

//     // Export data to Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);

//         // Apply date formats to worksheet columns
//         dateColumns.forEach(({ index, format }) => {
//             const colRef = XLSX.utils.encode_col(index);
//             worksheet[`!cols`] = worksheet[`!cols`] || [];
//             worksheet[`!cols`][index] = { z: format };
//         });

//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => {
//                                 let formattedCell = cell;
//                                 const expectedType = columnTypes[columnIndex];

//                                 if (expectedType === 'date' && cell) {
//                                     const dateColumn = dateColumns.find(col => col.index === columnIndex);
//                                     if (dateColumn) {
//                                         const parsedDate = new Date(cell);
//                                         if (!isNaN(parsedDate)) {
//                                             formattedCell = XLSX.SSF.format(dateColumn.format, parsedDate);
//                                         }
//                                     }
//                                 }

//                                 return (
//                                     <td key={columnIndex}>
//                                         <input
//                                             type="text"
//                                             value={formattedCell}
//                                             onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                         />
//                                     </td>
//                                 );
//                             })}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;



















// import React, { useState } from 'react';
// import * as XLSX from 'xlsx';
// import './excelEditor.css';

// function ExcelEditor() {
//     const [headers, setHeaders] = useState([]);
//     const [data, setData] = useState([]);
//     const [filteredData, setFilteredData] = useState([]);
//     const [filters, setFilters] = useState([]);
//     const [columnTypes, setColumnTypes] = useState([]);

//     // Handle file upload and read data
//     const handleFileUpload = (e) => {
//         const file = e.target.files[0];
//         const reader = new FileReader();

//         reader.onload = (event) => {
//             const binaryString = event.target.result;
//             const workbook = XLSX.read(binaryString, { type: 'binary' });
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//             // Extract headers (first row) and data rows
//             const headers = excelData[0];
//             const dataRows = excelData.slice(1);

//             // Determine column types and data rows formatting
//             const { types, formattedData } = processExcelData(dataRows, worksheet);

//             // Set headers, data, and filtered data
//             setHeaders(headers);
//             setData(formattedData);
//             setFilteredData(formattedData);

//             // Initialize filters with empty values
//             setFilters(Array(headers.length).fill(''));

//             // Set column types
//             setColumnTypes(types);
//         };

//         reader.readAsBinaryString(file);
//     };

//     // Process Excel data rows: determine column types and convert date/number columns
//     const processExcelData = (dataRows, worksheet) => {
//         const types = [];
//         const formattedData = [];

//         // Iterate through each column
//         for (let colIndex = 0; colIndex < dataRows[0].length; colIndex++) {
//             const cellFormat = worksheet[XLSX.utils.encode_cell({ c: colIndex, r: 1 })] ? worksheet[XLSX.utils.encode_cell({ c: colIndex, r: 1 })].z : null;
//             let columnType = 'string';

//             if (cellFormat && (cellFormat.includes('d') || cellFormat.includes('m') || cellFormat.includes('y'))) {
//                 columnType = 'date';
//             } else if (cellFormat && cellFormat.includes('0')) {
//                 columnType = 'number';
//             }

//             types.push(columnType);

//             // Format data based on column type and Excel format
//             formattedData.forEach((row, rowIndex) => {
//                 const cell = dataRows[rowIndex][colIndex];
//                 let formattedCell = cell;

//                 if (columnType === 'date') {
//                     formattedCell = formatExcelDate(cell);
//                 } else if (columnType === 'number') {
//                     formattedCell = formatExcelNumber(cell, cellFormat);
//                 }

//                 row[colIndex] = formattedCell;
//             });
//         }

//         // Convert data rows
//         dataRows.forEach((row, rowIndex) => {
//             formattedData[rowIndex] = row.map((cell, colIndex) => {
//                 const type = types[colIndex];
//                 if (type === 'date') {
//                     return formatExcelDate(cell);
//                 } else if (type === 'number') {
//                     return formatExcelNumber(cell);
//                 } else {
//                     return cell;
//                 }
//             });
//         });

//         return { types, formattedData };
//     };

//     // Format Excel date cell
//     const formatExcelDate = (cell) => {
//         const date = new Date(cell);
//         if (!isNaN(date.getTime())) {
//             return date.toISOString().slice(0, 10);
//         }
//         return cell;
//     };

//     // Format Excel number cell
//     const formatExcelNumber = (cell, format) => {
//         const number = parseFloat(cell);
//         if (!isNaN(number)) {
//             return number;
//         }
//         return cell;
//     };

//     // Apply filters to the data
//     const applyFilters = () => {
//         const filtered = data.filter(row =>
//             row.every((cell, index) => {
//                 if (filters[index] === '') {
//                     return true;
//                 }
//                 return String(cell).toLowerCase().includes(filters[index].toLowerCase());
//             })
//         );
//         setFilteredData(filtered);
//     };

//     // Handle cell change with data validation and update data state
//     const handleCellChange = (e, rowIndex, columnIndex) => {
//         let newValue = e.target.value;
//         const expectedType = columnTypes[columnIndex];

//         // Validate new value based on expected data type
//         if (!isValidValue(newValue, expectedType)) {
//             alert(`Invalid input! Expected a value of type ${expectedType}.`);
//             return;
//         }

//         // Format new value based on type
//         if (expectedType === 'date') {
//             const date = new Date(newValue);
//             if (!isNaN(date.getTime())) {
//                 newValue = date.toISOString().slice(0, 10);
//             } else {
//                 alert('Invalid date format! Please enter a valid date.');
//                 return;
//             }
//         } else if (expectedType === 'number') {
//             newValue = parseFloat(newValue);
//             if (isNaN(newValue)) {
//                 alert('Invalid number format! Please enter a valid number.');
//                 return;
//             }
//         }

//         // Update data state with the new value
//         const newData = [...data];
//         newData[rowIndex][columnIndex] = newValue;
//         setData(newData);
//         applyFilters();
//     };

//     // Validate value based on expected data type
//     const isValidValue = (value, type) => {
//         if (type === 'number') {
//             return !isNaN(parseFloat(value)) && isFinite(value);
//         } else if (type === 'date') {
//             const date = new Date(value);
//             return !isNaN(date.getTime());
//         } else {
//             return true;
//         }
//     };

//     // Add a new row with empty values
//     const addNewRow = () => {
//         const newRow = Array(headers.length).fill('');
//         setData([...data, newRow]);
//         setFilteredData([...filteredData, newRow]);
//     };

//     // Export data to an Excel file
//     const exportToExcel = (dataToExport, fileName) => {
//         const combinedData = [headers, ...dataToExport];
//         const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//         XLSX.writeFile(workbook, fileName);
//     };

//     return (
//         <div>
//             <h1>Excel Editor</h1>
//             <input type="file" onChange={handleFileUpload} />

//             {/* Render filter inputs above each column */}
//             <div className="filters">
//                 {filters.map((filter, index) => (
//                     <input
//                         key={index}
//                         type="text"
//                         placeholder={`Filter column ${index + 1}`}
//                         value={filter}
//                         onChange={(e) => {
//                             const newFilters = [...filters];
//                             newFilters[index] = e.target.value;
//                             setFilters(newFilters);
//                         }}
//                     />
//                 ))}
//                 {/* Button to apply filters */}
//                 <button onClick={applyFilters}>Apply Filters</button>
//             </div>

//             {/* Buttons to export data */}
//             <button onClick={() => exportToExcel(filteredData, 'filtered_data.xlsx')}>
//                 Export Filtered Data
//             </button>
//             <button onClick={() => exportToExcel(data, 'complete_data.xlsx')}>
//                 Export Complete Data
//             </button>

//             {/* Button to add a new row */}
//             <button onClick={addNewRow}>Add New Row</button>

//             {/* Render table */}
//             <table>
//                 <thead>
//                     <tr>
//                         {/* Render column headers */}
//                         {headers.map((header, index) => (
//                             <th key={index}>{header}</th>
//                         ))}
//                     </tr>
//                 </thead>
//                 <tbody>
//                     {/* Render filtered data */}
//                     {filteredData.map((row, rowIndex) => (
//                         <tr key={rowIndex}>
//                             {row.map((cell, columnIndex) => {
//                                 // Display date columns as date inputs
//                                 if (columnTypes[columnIndex] === 'date') {
//                                     return (
//                                         <td key={columnIndex}>
//                                             <input
//                                                 type="date"
//                                                 value={cell}
//                                                 onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                             />
//                                         </td>
//                                     );
//                                 } else if (columnTypes[columnIndex] === 'number') {
//                                     // Display number columns as number inputs
//                                     return (
//                                         <td key={columnIndex}>
//                                             <input
//                                                 type="number"
//                                                 value={cell}
//                                                 onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                             />
//                                         </td>
//                                     );
//                                 } else {
//                                     // Display other columns as text inputs
//                                     return (
//                                         <td key={columnIndex}>
//                                             <input
//                                                 type="text"
//                                                 value={cell}
//                                                 onChange={(e) => handleCellChange(e, rowIndex, columnIndex)}
//                                             />
//                                         </td>
//                                     );
//                                 }
//                             })}
//                         </tr>
//                     ))}
//                 </tbody>
//             </table>
//         </div>
//     );
// }

// export default ExcelEditor;
