import React from "react";

function DataTable({ sheetData, selectedSheetIndex, handleCellChange }) {
  return (
    <>
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
    </>
  );
}

export default DataTable;
