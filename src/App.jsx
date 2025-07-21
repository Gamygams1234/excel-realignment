import { useState } from "react";
import "./App.css";
import ExcelUploader from "./components/ExcelUploader";
import ColumnMapper from "./components/ColumnMapper";
import * as XLSX from "xlsx";

function App() {
  const [sourceFile, setSourceFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [sourceData, setSourceData] = useState(null);
  const [templateData, setTemplateData] = useState(null);
  const [templateHeaders, setTemplateHeaders] = useState([]);
  const [startColumn, setStartColumn] = useState("A");
  const [endColumn, setEndColumn] = useState(""); // New: end column for source data
  const [startRow, setStartRow] = useState(1);
  const [columnMapping, setColumnMapping] = useState({});
  const [sourceSheet, setSourceSheet] = useState("");
  const [templateSheet, setTemplateSheet] = useState("");
  const [templateHeaderRow, setTemplateHeaderRow] = useState(1);
  const [templateStartColumn, setTemplateStartColumn] = useState("A");
  const [readyToMap, setReadyToMap] = useState(false);
  const [customFilename, setCustomFilename] = useState(""); // New: custom filename
  
  // Store original workbook objects for formula preservation
  const [sourceWorkbook, setSourceWorkbook] = useState(null);
  const [templateWorkbook, setTemplateWorkbook] = useState(null);

  const extractHeadersFromSheet = (sheet, headerRow) => {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const headerRowIndex = headerRow - 1;
    const headersRow = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = { c, r: headerRowIndex };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = sheet[cellRef];
      headersRow.push(cell ? cell.v : "");
    }
    return headersRow;
  };

  const handleTemplateHeaderExtraction = (templateSheetObject) => {
    if (templateSheetObject) {
      const headers = extractHeadersFromSheet(templateSheetObject, templateHeaderRow);
      setTemplateHeaders(headers);
    }
  };

  // Validation logic for Continue button
  const isReadyToContinue = () => {
    return (
      sourceFile &&
      templateFile &&
      sourceSheet &&
      templateSheet &&
      startColumn &&
      startRow &&
      templateStartColumn &&
      templateHeaderRow
    );
  };

  const handleContinue = () => {
    if (isReadyToContinue()) {
      setReadyToMap(true);
    }
  };

  // Reset readyToMap when any critical data changes
  const resetMappingState = () => {
    setReadyToMap(false);
    setColumnMapping({});
  };

  // Reset when files or sheets change
  const handleSourceFileChange = (file, workbook) => {
    setSourceFile(file);
    setSourceWorkbook(workbook); // Store the workbook for formula preservation
    resetMappingState();
  };

  const handleTemplateFileChange = (file, workbook) => {
    setTemplateFile(file);
    setTemplateWorkbook(workbook); // Store the workbook for formula preservation
    resetMappingState();
  };

  const handleSourceSheetChange = (sheet) => {
    setSourceSheet(sheet);
    resetMappingState();
  };

  const handleTemplateSheetChange = (sheet) => {
    setTemplateSheet(sheet);
    resetMappingState();
  };

  return (
    <div className="container">
      <h1>Excel Sheet Realigner</h1>
      <div className="upload-section">
        <ExcelUploader
          file={sourceFile}
          onFileChange={handleSourceFileChange}
          onDataLoaded={setSourceData}
          title="Source File"
          onColumnChange={setStartColumn}
          onEndColumnChange={setEndColumn}
          onRowChange={setStartRow}
          startColumn={startColumn}
          endColumn={endColumn}
          startRow={startRow}
          onSheetChange={handleSourceSheetChange}
        />
        <ExcelUploader
          file={templateFile}
          onFileChange={handleTemplateFileChange}
          onDataLoaded={setTemplateData}
          title="Template File"
          onSheetChange={handleTemplateSheetChange}
          onHeaderRowChange={setTemplateHeaderRow}
          onTemplateColumnChange={setTemplateStartColumn}
          onSheetParsed={handleTemplateHeaderExtraction}
        />
      </div>

      {/* Continue Button Section */}
      <div style={{ textAlign: "center", margin: "2rem 0" }}>
        <button
          onClick={handleContinue}
          disabled={!isReadyToContinue()}
          style={{
            padding: "12px 24px",
            fontSize: "16px",
            fontWeight: "bold",
            backgroundColor: isReadyToContinue() ? "#1976d2" : "#ccc",
            color: "white",
            border: "none",
            borderRadius: "8px",
            cursor: isReadyToContinue() ? "pointer" : "not-allowed",
            transition: "all 0.3s ease",
          }}
        >
          Continue to Column Mapping
        </button>

        {/* Validation Feedback */}
        {!isReadyToContinue() && (
          <div style={{ marginTop: "1rem", color: "#666", fontSize: "14px" }}>
            <p>Please complete all required fields:</p>
            <ul style={{ textAlign: "left", display: "inline-block", margin: 0 }}>
              {!sourceFile && <li>Upload source file</li>}
              {!templateFile && <li>Upload template file</li>}
              {!sourceSheet && <li>Select source sheet</li>}
              {!templateSheet && <li>Select template sheet</li>}
              {!startColumn && <li>Set source start column</li>}
              {!startRow && <li>Set source start row</li>}
              {!templateStartColumn && <li>Set template start column</li>}
              {!templateHeaderRow && <li>Set template header row</li>}
            </ul>
          </div>
        )}
      </div>

      {/* Column Mapping Section - Only show when ready */}
      {sourceData && templateHeaders.length > 0 && (
        <ColumnMapper
          sourceData={sourceData}
          templateHeaders={templateHeaders}
          startColumn={startColumn}
          endColumn={endColumn}
          startRow={startRow}
          columnMapping={columnMapping}
          onMappingChange={setColumnMapping}
          sourceWorkbook={sourceWorkbook}
          sourceSheet={sourceSheet}
          templateWorkbook={templateWorkbook}
          templateSheet={templateSheet}
          customFilename={customFilename}
          onFilenameChange={setCustomFilename}
        />
      )}
    </div>
  );
}

export default App;
