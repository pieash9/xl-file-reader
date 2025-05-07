"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

const ExcelReader = () => {
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [columnNames, setColumnNames] = useState([]);
  const [selectedColumn, setSelectedColumn] = useState("");
  const [columnData, setColumnData] = useState([]);
  const [workbookData, setWorkbookData] = useState(null);

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return alert("Please upload a file");

    const reader = new FileReader();
    reader.onload = (event) => {
      const arrayBuffer = event.target.result;
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: "array" });

      setWorkbookData(workbook);
      setSheetNames(workbook.SheetNames);
      setSelectedSheet(workbook.SheetNames[0]);
      handleSheetSelection(workbook.SheetNames[0], workbook);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleSheetSelection = (sheetName, workbook) => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    if (jsonData.length > 0) {
      const headers = Object.keys(jsonData[0]);
      setColumnNames(headers);
      setSelectedColumn(headers[0]);
      extractColumnData(headers[0], jsonData);
    } else {
      setColumnNames([]);
      setColumnData([]);
    }
  };

  const handleSheetChange = (e) => {
    const newSheet = e.target.value;
    setSelectedSheet(newSheet);
    if (workbookData) {
      handleSheetSelection(newSheet, workbookData);
    }
  };

  const handleColumnChange = (e) => {
    const column = e.target.value;
    setSelectedColumn(column);

    const worksheet = workbookData.Sheets[selectedSheet];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    extractColumnData(column, jsonData);
  };

  const extractColumnData = (columnName, data) => {
    const values = data
      .map((row) => row[columnName])
      .filter((val) => val !== undefined && val !== null);
    setColumnData(values);
  };

  return (
    <div className="p-4">
      <label htmlFor="input-file" className="block mb-2 text-sm">
        Upload Excel File
      </label>
      <input
        id="input-file"
        className="mb-4 block w-fit text-sm border border-gray-300 rounded-lg cursor-pointer bg-gray-500 p-2"
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
      />

      {sheetNames.length > 0 && (
        <div className="mb-4">
          <label className="block mb-1 text-sm font-medium">
            Select Sheet:
          </label>
          <select
            value={selectedSheet}
            onChange={handleSheetChange}
            className="p-1 border rounded "
          >
            {sheetNames.map((name) => (
              <option className="text-black" key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>
      )}

      {columnNames.length > 0 && (
        <>
          <label className="block mb-2 text-sm font-medium">
            Select Column:
            <select
              value={selectedColumn}
              onChange={handleColumnChange}
              className="ml-2 p-1 border rounded "
            >
              {columnNames.map((col) => (
                <option className="text-black" key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>
          </label>

          <h2 className="mt-4 text-xl font-bold">
            Data from "{selectedColumn}" column:
          </h2>
          <ul className="mt-2 list-disc pl-5">
            {columnData.map((item, index) => (
              <li key={index}>{item}</li>
            ))}
          </ul>
        </>
      )}
    </div>
  );
};

export default ExcelReader;
