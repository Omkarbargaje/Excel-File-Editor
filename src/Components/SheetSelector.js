import React from "react";

function SheetSelector({ sheetData, handleSheetChange, activeSheetIndex }) {
  return (
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
  );
}

export default SheetSelector;
