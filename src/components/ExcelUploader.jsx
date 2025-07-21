import React, { useState } from 'react';
import { Button, TextField, Box, Typography, FormControl, InputLabel, Select, MenuItem } from '@mui/material';
import { Upload as UploadIcon } from '@mui/icons-material';
import * as XLSX from 'xlsx';

const ExcelUploader = ({
  file,
  onFileChange,
  onDataLoaded,
  title,
  onColumnChange,
  onEndColumnChange,
  onRowChange,
  startColumn,
  endColumn,
  startRow,
  onSheetChange,
  onHeaderRowChange,
  onTemplateColumnChange,
  onSheetParsed
}) => {
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [workbook, setWorkbook] = useState(null);
  const [headerRow, setHeaderRow] = useState(1);
  const [templateStartColumn, setTemplateStartColumn] = useState('A');

  const handleFileChange = async (event) => {
    const file = event.target.files[0];
    if (file) {
      const { workbook: wb, sheetNames: sheets } = await parseExcelFile(file);
      setWorkbook(wb);
      setSheetNames(sheets);
      setSelectedSheet(sheets[0]);

      // Pass both file and workbook to parent
      onFileChange(file, wb);
      
      const data = parseSheetData(wb, sheets[0]);
      onDataLoaded(data);
      if (onSheetChange) onSheetChange(sheets[0]);

      if (title === "Template File" && onSheetParsed) {
        const worksheet = wb.Sheets[sheets[0]];
        onSheetParsed(worksheet);
      }
    }
  };

  const parseExcelFile = async (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetNames = workbook.SheetNames;
        resolve({ workbook, sheetNames });
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    });
  };

  const parseSheetData = (wb, sheetName) => {
    const worksheet = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  };

  const parseTemplateSheetData = (wb, sheetName, headerRowNum, startCol) => {
    const worksheet = wb.Sheets[sheetName];
    const fullData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (fullData.length < headerRowNum) return [];
    
    // Get only the header row
    const headerRowData = fullData[headerRowNum - 1];
    const startColIndex = startCol.charCodeAt(0) - 'A'.charCodeAt(0);
    const extractedHeaders = headerRowData.slice(startColIndex);
    
    // Return only the header row as a single-row array
    return [extractedHeaders];
  };

  const handleSheetChange = (event) => {
    const newSheet = event.target.value;
    setSelectedSheet(newSheet);
    if (workbook) {
      let data;
      if (title === "Template File") {
        data = parseTemplateSheetData(workbook, newSheet, headerRow, templateStartColumn);
      } else {
        data = parseSheetData(workbook, newSheet);
      }
      onDataLoaded(data);
      if (onSheetChange) onSheetChange(newSheet);
      
      // For Template File: trigger header extraction
      if (title === "Template File" && onSheetParsed) {
        const worksheet = workbook.Sheets[newSheet];
        onSheetParsed(worksheet);
      }
    }
  };

  const handleHeaderRowChange = (event) => {
    const newHeaderRow = Number(event.target.value);
    setHeaderRow(newHeaderRow);
    if (onHeaderRowChange) onHeaderRowChange(newHeaderRow);
    
    // Re-parse data with new header row for template files
    if (title === "Template File" && workbook && selectedSheet) {
      const data = parseTemplateSheetData(workbook, selectedSheet, newHeaderRow, templateStartColumn);
      onDataLoaded(data);
      
      // Re-trigger header extraction
      if (onSheetParsed) {
        const worksheet = workbook.Sheets[selectedSheet];
        onSheetParsed(worksheet);
      }
    }
  };

  const handleTemplateColumnChange = (event) => {
    const newStartColumn = event.target.value.toUpperCase();
    setTemplateStartColumn(newStartColumn);
    if (onTemplateColumnChange) onTemplateColumnChange(newStartColumn);
    
    // Re-parse data with new start column for template files
    if (title === "Template File" && workbook && selectedSheet) {
      const data = parseTemplateSheetData(workbook, selectedSheet, headerRow, newStartColumn);
      onDataLoaded(data);
      
      // Re-trigger header extraction
      if (onSheetParsed) {
        const worksheet = workbook.Sheets[selectedSheet];
        onSheetParsed(worksheet);
      }
    }
  };

  return (
    <Box sx={{ p: 2, border: '1px solid #ccc', borderRadius: 2, mb: 2 }}>
      <Typography variant="h6" gutterBottom>
        {title}
      </Typography>

      <input
        accept=".xlsx,.xls"
        style={{ display: 'none' }}
        id={`excel-upload-${(title || '').toLowerCase().replace(' ', '-')}`}
        type="file"
        onChange={handleFileChange}
      />

      <label htmlFor={`excel-upload-${(title || '').toLowerCase().replace(' ', '-')}`}>
        <Button
          variant="contained"
          component="span"
          startIcon={<UploadIcon />}
          fullWidth
        >
          {file ? file.name : 'Upload Excel File'}
        </Button>
      </label>

      {/* Sheet Selection Dropdown */}
      {sheetNames.length > 0 && (
        <Box sx={{ mt: 2 }}>
          <FormControl fullWidth>
            <InputLabel>Select Sheet</InputLabel>
            <Select
              value={selectedSheet}
              onChange={handleSheetChange}
              label="Select Sheet"
            >
              {sheetNames.map((sheetName) => (
                <MenuItem key={sheetName} value={sheetName}>
                  {sheetName}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </Box>
      )}

      {/* Source File specific controls */}
      {title === "Source File" && sheetNames.length > 0 && (
        <Box sx={{ mt: 2 }}>
          <Box sx={{ display: 'flex', gap: 2, mb: 2 }}>
            <TextField
              label="Start Column"
              value={startColumn || ''}
              onChange={(e) => onColumnChange((e.target.value || '').toUpperCase())}
              sx={{ flex: 1 }}
              helperText="First column to extract"
            />
            <TextField
              label="End Column (Optional)"
              value={endColumn || ''}
              onChange={(e) => onEndColumnChange && onEndColumnChange((e.target.value || '').toUpperCase())}
              sx={{ flex: 1 }}
              helperText="Last column to extract (leave blank for all)"
            />
            <TextField
              label="Start Row"
              type="number"
              value={startRow || ''}
              onChange={(e) => onRowChange(Number(e.target.value))}
              sx={{ flex: 1 }}
              helperText="Row containing headers"
            />
          </Box>
        </Box>
      )}

      {/* Template File specific controls */}
      {title === "Template File" && sheetNames.length > 0 && (
        <Box sx={{ mt: 2 }}>
          <Box sx={{ display: 'flex', gap: 2, mb: 2 }}>
            <TextField
              label="Start Column (Optional)"
              value={templateStartColumn}
              onChange={handleTemplateColumnChange}
              sx={{ flex: 1 }}
              helperText="Column to start extracting headers from"
            />
            <TextField
              label="Header Row (Required)"
              type="number"
              value={headerRow}
              onChange={handleHeaderRowChange}
              sx={{ flex: 1 }}
              helperText="Row number containing the headers"
            />
          </Box>
        </Box>
      )}
    </Box>
  );
};

export default ExcelUploader;
