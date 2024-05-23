import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

import React from "react";

function ExportButtons({
  sheetData,
  originalData,
  selectedSheetIndex,
  workbook,
}) {
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

  return (
    <>
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
    </>
  );
}

export default ExportButtons;
