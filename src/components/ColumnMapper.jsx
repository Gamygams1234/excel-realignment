import React, { useEffect } from 'react';
import { Box, Typography, FormControl, InputLabel, Select, MenuItem, Button, TextField } from '@mui/material';
import { Download as DownloadIcon, AutoFixHigh as AutoMapIcon } from '@mui/icons-material';
import * as XLSX from 'xlsx';

const ColumnMapper = ({ sourceData, templateHeaders, startColumn, endColumn, startRow, columnMapping, onMappingChange, sourceWorkbook, sourceSheet, templateWorkbook, templateSheet, customFilename, onFilenameChange }) => {
  const getSourceHeaders = () => {
    if (!sourceData || !Array.isArray(sourceData)) return [];
    const startColIndex = getColIndex(startColumn);
    const headers = sourceData[startRow - 1];
    if (!headers || !Array.isArray(headers)) return [];
    
    let extractedHeaders;
    if (endColumn && endColumn.trim()) {
      // If end column is specified, extract only the range
      const endColIndex = getColIndex(endColumn);
      extractedHeaders = headers.slice(startColIndex, endColIndex + 1);
    } else {
      // If no end column, extract from start column to end
      extractedHeaders = headers.slice(startColIndex);
    }
    
    return extractedHeaders.filter(header => header !== null && header !== undefined);
  };

  const getColIndex = (col) => {
    return col.charCodeAt(0) - 'A'.charCodeAt(0);
  };

  const handleMappingChange = (sourceCol, event) => {
    const newMapping = { ...columnMapping };
    newMapping[sourceCol] = event.target.value;
    onMappingChange(newMapping);
  };

  const autoMapColumns = () => {
    const sourceHeaders = getSourceHeaders();
    const newMapping = {};

    sourceHeaders.forEach(sourceHeader => {
      if (sourceHeader && sourceHeader.toString && sourceHeader.toString().trim()) {
        const matchedTemplate = templateHeaders.find(templateHeader => 
          templateHeader && templateHeader.toString && 
          templateHeader.toString().toLowerCase().trim() === sourceHeader.toString().toLowerCase().trim()
        );
        if (matchedTemplate) {
          newMapping[sourceHeader] = matchedTemplate;
        }
      }
    });

    onMappingChange(newMapping);
  };

  useEffect(() => {
    if (sourceData && templateHeaders && Array.isArray(templateHeaders) && templateHeaders.length > 0 && Object.keys(columnMapping).length === 0) {
      autoMapColumns();
    }
  }, [sourceData, templateHeaders, startColumn, startRow]);

  const downloadRealignedFile = () => {
    try {
      if (!templateHeaders || templateHeaders.length === 0) {
        throw new Error('No template headers found. Please check your template file.');
      }

      // Get the full source headers row (before slicing)
      const fullSourceHeaders = sourceData[startRow - 1];
      if (!fullSourceHeaders || !Array.isArray(fullSourceHeaders)) {
        throw new Error('No source headers found. Please check your source file.');
      }

      // Get source data starting from the row after headers
      const sourceDataWithoutHeaders = sourceData.slice(startRow);

      if (sourceDataWithoutHeaders.length === 0) {
        throw new Error('No source data found. Please check your start row setting.');
      }

      console.log('Template Headers:', templateHeaders);
      console.log('Full Source Headers:', fullSourceHeaders);
      console.log('Column Mapping:', columnMapping);

      const realignedData = [
        templateHeaders, // First row: template headers define the structure
        ...sourceDataWithoutHeaders.map(sourceRow => {
          const newRow = templateHeaders.map(templateHeader => {
            // Find which source header is mapped to this template header
            const mappedSourceHeader = Object.entries(columnMapping)
              .find(([_, mappedTemplateHeader]) => mappedTemplateHeader === templateHeader)?.[0];
            
            if (!mappedSourceHeader) {
              // No mapping found - leave blank
              return '';
            }

            // Find the actual column index in the FULL source headers (not filtered by end column)
            const sourceColIndex = fullSourceHeaders.indexOf(mappedSourceHeader);
            
            if (sourceColIndex === -1) {
              // Source header not found - leave blank
              return '';
            }

            // Extract the data from the source row at the correct column index
            return sourceRow[sourceColIndex] || '';
          });
          return newRow;
        })
      ];

      console.log('Realigned Data Sample:', realignedData.slice(0, 3)); // Log first 3 rows for debugging

      const ws = XLSX.utils.aoa_to_sheet(realignedData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Realigned_Data");

      const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
      const baseFilename = customFilename && customFilename.trim() ? customFilename.trim() : 'realigned';
      const filename = `${baseFilename}_${timestamp}.xlsx`;

      XLSX.writeFile(wb, filename);
      alert(`File downloaded successfully as ${filename}\n\nStructure: ${templateHeaders.length} columns, ${sourceDataWithoutHeaders.length} data rows`);
    } catch (error) {
      console.error('Download error:', error);
      alert(`Error: ${error.message}`);
    }
  };

  const downloadWithFormulas = () => {
    try {
      if (!templateHeaders || templateHeaders.length === 0) {
        throw new Error('No template headers found. Please check your template file.');
      }

      if (!sourceWorkbook || !sourceSheet) {
        throw new Error('Source workbook not available. Please re-upload your source file.');
      }

      console.log('Starting formula preservation download...');
      console.log('Source Workbook:', sourceWorkbook);
      console.log('Source Sheet:', sourceSheet);
      console.log('Template Headers:', templateHeaders);
      console.log('Column Mapping:', columnMapping);

      // Get the source worksheet
      const sourceWorksheet = sourceWorkbook.Sheets[sourceSheet];
      if (!sourceWorksheet) {
        throw new Error(`Source sheet '${sourceSheet}' not found in workbook.`);
      }

      // Create a copy of the source workbook
      const newWorkbook = XLSX.utils.book_new();
      
      // Copy all existing sheets from source workbook
      Object.keys(sourceWorkbook.Sheets).forEach(sheetName => {
        const originalSheet = sourceWorkbook.Sheets[sheetName];
        const copiedSheet = XLSX.utils.sheet_to_json(originalSheet, { header: 1, raw: false, defval: '' });
        const newSheet = XLSX.utils.aoa_to_sheet(copiedSheet);
        
        // Copy cell formulas and formatting
        if (originalSheet['!ref']) {
          const range = XLSX.utils.decode_range(originalSheet['!ref']);
          for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
              const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
              const originalCell = originalSheet[cellAddress];
              if (originalCell) {
                if (!newSheet[cellAddress]) newSheet[cellAddress] = {};
                // Copy formula if it exists
                if (originalCell.f) {
                  newSheet[cellAddress].f = originalCell.f;
                }
                // Copy other cell properties
                if (originalCell.t) newSheet[cellAddress].t = originalCell.t;
                if (originalCell.v !== undefined) newSheet[cellAddress].v = originalCell.v;
                if (originalCell.w) newSheet[cellAddress].w = originalCell.w;
              }
            }
          }
          newSheet['!ref'] = originalSheet['!ref'];
        }
        
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
      });

      // Create the reorganized sheet with formulas preserved
      const reorganizedSheet = createReorganizedSheetWithFormulas(sourceWorksheet);
      XLSX.utils.book_append_sheet(newWorkbook, reorganizedSheet, 'Realigned_Data');

      // Download the workbook with all sheets including the reorganized one
      const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
      const baseFilename = customFilename && customFilename.trim() ? customFilename.trim() : 'realigned_with_formulas';
      const filename = `${baseFilename}_${timestamp}.xlsx`;

      XLSX.writeFile(newWorkbook, filename);
      alert(`File downloaded successfully as ${filename}\n\nIncludes:\n- All original sheets with formulas preserved\n- New 'Realigned_Data' sheet with reorganized structure\n- ${templateHeaders.length} columns, formulas maintained where possible`);
      
    } catch (error) {
      console.error('Formula download error:', error);
      alert(`Error: ${error.message}`);
    }
  };

  const createReorganizedSheetWithFormulas = (sourceWorksheet) => {
    try {
      // Get the full source headers row
      const fullSourceHeaders = sourceData[startRow - 1];
      if (!fullSourceHeaders || !Array.isArray(fullSourceHeaders)) {
        throw new Error('No source headers found.');
      }

      // Create new worksheet
      const newSheet = {};
      
      // Add template headers at the same row as source headers (startRow - 1 in 0-based indexing)
      const headerRowIndex = startRow - 1;
      templateHeaders.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: colIndex });
        newSheet[cellAddress] = { t: 's', v: header };
      });

      // Get source data rows (excluding header)
      const sourceDataRows = sourceData.slice(startRow);
      
      // Process each data row - maintain same row numbers as source
      sourceDataRows.forEach((sourceRow, rowIndex) => {
        const newRowIndex = startRow + rowIndex; // Keep same row numbers as source
        
        templateHeaders.forEach((templateHeader, colIndex) => {
          // Find which source header is mapped to this template header
          const mappedSourceHeader = Object.entries(columnMapping)
            .find(([_, mappedTemplateHeader]) => mappedTemplateHeader === templateHeader)?.[0];
          
          if (!mappedSourceHeader) {
            // No mapping - leave blank
            const cellAddress = XLSX.utils.encode_cell({ r: newRowIndex, c: colIndex });
            newSheet[cellAddress] = { t: 's', v: '' };
            return;
          }

          // Find the source column index in the FULL source headers (not filtered by end column)
          const sourceColIndex = fullSourceHeaders.indexOf(mappedSourceHeader);
          if (sourceColIndex === -1) {
            // Source header not found - leave blank
            const cellAddress = XLSX.utils.encode_cell({ r: newRowIndex, c: colIndex });
            newSheet[cellAddress] = { t: 's', v: '' };
            return;
          }

          // Get the original cell from source worksheet
          const sourceCellAddress = XLSX.utils.encode_cell({ r: startRow + rowIndex, c: sourceColIndex });
          const sourceCell = sourceWorksheet[sourceCellAddress];
          
          const newCellAddress = XLSX.utils.encode_cell({ r: newRowIndex, c: colIndex });
          
          if (sourceCell) {
            // Copy the cell with formula if it exists
            newSheet[newCellAddress] = { ...sourceCell };
            
            // If it's a formula, we might need to adjust references
            // For now, we'll copy as-is and let Excel handle relative references
            if (sourceCell.f) {
              console.log(`Copying formula from ${sourceCellAddress} to ${newCellAddress}: ${sourceCell.f}`);
            }
          } else {
            // No source cell - use the parsed value or leave blank
            const value = sourceRow[sourceColIndex] || '';
            newSheet[newCellAddress] = { t: 's', v: value };
          }
        });
      });

      // Set the sheet range - from header row to last data row
      const minRow = headerRowIndex; // Start from header row
      const maxRow = startRow + sourceDataRows.length - 1; // Last data row
      const maxCol = templateHeaders.length - 1;
      newSheet['!ref'] = XLSX.utils.encode_range({ s: { r: minRow, c: 0 }, e: { r: maxRow, c: maxCol } });
      
      return newSheet;
      
    } catch (error) {
      console.error('Error creating reorganized sheet:', error);
      throw error;
    }
  };

  return (
    <Box sx={{ p: 2, border: '1px solid #ccc', borderRadius: 2, mt: 2 }}>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
        <Typography variant="h6">Column Mapping</Typography>
        <Button variant="outlined" startIcon={<AutoMapIcon />} onClick={autoMapColumns} size="small">
          Auto-Map
        </Button>
      </Box>

      <Typography variant="subtitle2" sx={{ mb: 2 }}>Map source columns to template columns:</Typography>

      <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
        {getSourceHeaders().map((header, index) => {
          const isAutoMapped = columnMapping[header] && 
            templateHeaders.some(th => 
              th && header && 
              th.toString().toLowerCase().trim() === header.toString().toLowerCase().trim()
            );

          return (
            <Box key={index} sx={{ 
              display: 'flex', 
              alignItems: 'center', 
              gap: 2, 
              p: 1,
              borderRadius: 1,
              bgcolor: isAutoMapped ? 'success.light' : 'transparent',
              border: isAutoMapped ? '1px solid' : 'none',
              borderColor: isAutoMapped ? 'success.main' : 'transparent'
            }}>
              <Box sx={{ minWidth: 150 }}>
                <Typography sx={{ fontWeight: isAutoMapped ? 'bold' : 'normal' }}>
                  {header}
                  {isAutoMapped && (
                    <Typography component="span" sx={{ ml: 1, fontSize: '0.75rem', color: 'success.dark' }}>
                      ✓ Auto-mapped
                    </Typography>
                  )}
                </Typography>
              </Box>
              <FormControl fullWidth>
                <InputLabel>Map to</InputLabel>
                <Select
                  value={columnMapping[header] || ''}
                  onChange={(e) => handleMappingChange(header, e)}
                  label="Map to"
                >
                  <MenuItem value="">
                    <em>No mapping</em>
                  </MenuItem>
                  {(templateHeaders || []).map((templateHeader, idx) => (
                    <MenuItem key={idx} value={templateHeader || ''}>
                      {templateHeader || ''}
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Box>
          );
        })}
      </Box>

      {/* Mapping Summary */}
      {Object.keys(columnMapping).length > 0 && (
        <Box sx={{ mt: 2, p: 1, bgcolor: 'grey.50', borderRadius: 1 }}>
          <Typography variant="caption" color="text.secondary">
            📊 Mapping Summary: {Object.values(columnMapping).filter(v => v).length} of {getSourceHeaders().length} source columns mapped
          </Typography>
        </Box>
      )}

      {/* Custom Filename Input */}
      <Box sx={{ mt: 2 }}>
        <TextField
          label="Custom Filename (Optional)"
          value={customFilename || ''}
          onChange={(e) => onFilenameChange && onFilenameChange(e.target.value)}
          fullWidth
          helperText="Leave blank for auto-generated filename with timestamp"
          placeholder="e.g., my_realigned_data"
        />
      </Box>

      <Box sx={{ display: 'flex', gap: 2, mt: 2 }}>
        <Button
          variant="contained"
          startIcon={<DownloadIcon />}
          onClick={downloadRealignedFile}
          disabled={Object.values(columnMapping).every(v => !v)}
          size="large"
          sx={{ flex: 1 }}
        >
          Download (Values Only)
        </Button>
        
        <Button
          variant="outlined"
          startIcon={<DownloadIcon />}
          onClick={downloadWithFormulas}
          disabled={Object.values(columnMapping).every(v => !v)}
          size="large"
          sx={{ flex: 1 }}
        >
          Download (Keep Formulas)
        </Button>
      </Box>

      {Object.values(columnMapping).some(v => v) && (
        <Box sx={{ mt: 2, p: 2, bgcolor: 'grey.50', borderRadius: 1 }}>
          <Typography variant="subtitle2" sx={{ mb: 1, fontWeight: 'bold' }}>
            Download Options:
          </Typography>
          <Typography variant="caption" display="block" sx={{ mb: 1 }}>
            📊 <strong>Values Only:</strong> Downloads reorganized data as values (current functionality)
          </Typography>
          <Typography variant="caption" display="block">
            📝 <strong>Keep Formulas:</strong> Preserves Excel formulas, formatting, and adds reorganized sheet to original workbook
          </Typography>
        </Box>
      )}
    </Box>
  );
};

export default ColumnMapper;
