import * as XLSX from "xlsx";
import React, { useState } from "react";
import "./excelEditor.css";

function FileUploader({ setWorkbook, workbook }) {
  const [selectedFile, setSelectedFile] = useState(null);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    setSelectedFile(file);
  };

  const handleFileUpload = () => {
    if (!selectedFile) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const binaryString = event.target.result;
      const wb = XLSX.read(binaryString, { type: "binary" });
      setWorkbook(wb);
    };
    reader.readAsBinaryString(selectedFile);
  };

  return (
    <>
      {!workbook && (
        <label className="fileUploaderLabel">Please select a file</label>
      )}
      <div className="file-input-container">
        <label htmlFor="file-upload" id="custom-file-upload">
          Choose File
        </label>
        <input
          id="file-upload"
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileChange}
        />
        {selectedFile && <span className="file-name">{selectedFile.name}</span>}
        <button onClick={handleFileUpload} className="upload-btn">
          Submit
        </button>
      </div>
    </>
  );
}

export default FileUploader;
