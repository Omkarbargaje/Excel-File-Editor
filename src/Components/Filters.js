import React from "react";

function Filters({
  filters,
  globalSearch,
  selectedSheetIndex,
  originalData,
  setSheetData,
  setFilters,
  sheetData,
  setGlobalSearch,
}) {
  const applyGlobalFilters = (handleSheetChange) => {
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

  const applyFilters = (handleSheetChange) => {
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
  return (
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
  );
}

export default Filters;
