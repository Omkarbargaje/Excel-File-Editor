import React from "react";

function AddNewRow({
  sheetData,
  selectedSheetIndex,
  setSheetData,
  setOriginalData,
}) {
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
  return <button onClick={addNewRow}>Add New Row</button>;
}

export default AddNewRow;
