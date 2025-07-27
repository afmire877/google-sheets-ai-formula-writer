function getSelectedRange() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getActiveRange();
}

function getSelectedValues() {
  const range = getSelectedRange();
  if (!range) return null;
  return range.getValues();
}


function getSelectedHeaders() {
  const values = getSelectedValues();
  if (!values || values.length === 0) return null;
  return values[0];
}


function summarizeSelectedColumns() {
  const values = getSelectedValues();
  if (!values || values.length < 2) return null;

  const sums = [];
  const numCols = values[0].length;


  // Initialize sums array
  for (let c = 0; c < numCols; c++) {
    sums[c] = 0;
  }

  // Start from row 1 to skip header row (index 0)
  for (let r = 1; r < values.length; r++) {
    for (let c = 0; c < numCols; c++) {
      const val = values[r][c];
      if (typeof val === 'number') {
        sums[c] += val;
      }
    }
  }

  return sums;
}

// New functions for enhanced modal
function analyzeSelectedData() {
  try {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) return { error: "No range selected" };
    
    const values = range.getValues();
    if (!values || values.length === 0) return { error: "No data in selected range" };
    
    const analysis = {
      rangeAddress: range.getA1Notation(),
      rowCount: values.length,
      colCount: values[0].length,
      hasHeaders: false,
      columns: []
    };
    
    // Analyze each column
    for (let col = 0; col < values[0].length; col++) {
      const columnData = values.map(row => row[col]).filter(cell => cell !== "");
      const columnAnalysis = {
        index: col,
        header: values[0][col],
        dataType: detectDataType(columnData.slice(1)),
        sampleValues: columnData.slice(1, 4),
        uniqueCount: new Set(columnData.slice(1)).size,
        hasNumbers: columnData.slice(1).some(val => typeof val === 'number'),
        hasText: columnData.slice(1).some(val => typeof val === 'string'),
        hasDates: columnData.slice(1).some(val => val instanceof Date)
      };
      analysis.columns.push(columnAnalysis);
    }
    
    // Check if first row looks like headers
    analysis.hasHeaders = values[0].every(cell => typeof cell === 'string' && cell.length > 0);
    
    return analysis;
  } catch (error) {
    return { error: error.toString() };
  }
}

function generateAIFormula(promptText) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
      return { success: false, error: "No range selected for context" };
    }
    
    const values = range.getValues();
    const analysis = analyzeSelectedData();
    
    // Create enhanced prompt with data context
    const contextualPrompt = createContextualPrompt(promptText, analysis, values);
    
    // Call OpenAI with the enhanced prompt
    const response = callOpenAIForFormula(contextualPrompt);
    
    if (!response) {
      return { success: false, error: "No response from AI" };
    }
    
    // Try to extract formula from response
    const formula = extractFormulaFromResponse(response);
    
    if (formula) {
      // Find appropriate cell to insert formula
      const targetCell = findTargetCell(range);
      
      try {
        const sheet = SpreadsheetApp.getActiveSheet();
        const cell = sheet.getRange(targetCell);
        cell.setFormula(formula);
        
        return {
          success: true,
          formula: formula,
          cellAddress: targetCell,
          inserted: true,
          explanation: response
        };
      } catch (insertError) {
        return {
          success: true,
          formula: formula,
          cellAddress: targetCell,
          inserted: false,
          error: "Formula generated but insertion failed: " + insertError.toString(),
          explanation: response
        };
      }
    } else {
      return {
        success: false,
        error: "Could not extract formula from AI response",
        response: response
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function createContextualPrompt(userRequest, analysis, values) {
  let prompt = `USER REQUEST: ${userRequest}\n\n`;
  
  // Add data context
  prompt += `SELECTED DATA CONTEXT:\n`;
  prompt += `Range: ${analysis.rangeAddress}\n`;
  prompt += `Size: ${analysis.rowCount} rows × ${analysis.colCount} columns\n`;
  
  // Add column information with letters
  if (analysis.columns && analysis.columns.length > 0) {
    prompt += `Columns:\n`;
    analysis.columns.forEach((col, index) => {
      const columnLetter = String.fromCharCode(65 + index); // A, B, C, etc.
      prompt += `  ${columnLetter}: "${col.header}" (${col.dataType})\n`;
    });
  }
  
  // Add sample data with proper formatting
  prompt += `\nSAMPLE DATA:\n`;
  if (analysis.hasHeaders && values.length > 0) {
    // Show headers
    const headers = values[0].map((header, i) => `${String.fromCharCode(65 + i)}:${header}`);
    prompt += `Headers: ${headers.join(', ')}\n`;
    
    // Show data rows
    values.slice(1, Math.min(4, values.length)).forEach((row, i) => {
      prompt += `Row ${i + 2}: ${row.join(' | ')}\n`;
    });
  } else {
    values.slice(0, 3).forEach((row, i) => {
      prompt += `Row ${i + 1}: ${row.join(' | ')}\n`;
    });
  }
  
  // Add specific instructions
  prompt += `\nINSTRUCTIONS:\n`;
  prompt += `- Generate a formula that works with the selected range ${analysis.rangeAddress}\n`;
  prompt += `- Use proper column references (A, B, C, etc.)\n`;
  prompt += `- Consider the data types when choosing functions\n`;
  prompt += `- Return only the formula starting with =\n`;
  
  return prompt;
}

function extractFormulaFromResponse(response) {
  // Multiple regex patterns to catch different formula formats
  const patterns = [
    /^=.+$/m,                    // Formula at start of line
    /(?:^|\n)=.+$/gm,           // Formula after newline
    /(?:Formula|formula):\s*=.+$/gm,  // "Formula: =..." format
    /```\s*=.+\s*```/gm,        // Code blocks with formulas
    /`=.+`/gm                   // Inline code with formulas
  ];
  
  for (const pattern of patterns) {
    const matches = response.match(pattern);
    if (matches && matches.length > 0) {
      let formula = matches[0];
      
      // Clean up the formula
      formula = formula.replace(/^(?:Formula|formula):\s*/i, '');
      formula = formula.replace(/```/g, '');
      formula = formula.replace(/`/g, '');
      formula = formula.trim();
      
      // Ensure it starts with =
      if (formula.startsWith('=')) {
        return formula;
      }
    }
  }
  
  // If no pattern matches, check if the entire response is just a formula
  const cleanResponse = response.trim();
  if (cleanResponse.startsWith('=') && !cleanResponse.includes('\n')) {
    return cleanResponse;
  }
  
  return null;
}

function findTargetCell(selectedRange) {
  const sheet = selectedRange.getSheet();
  const lastCol = selectedRange.getLastColumn();
  const lastRow = selectedRange.getLastRow();
  const firstRow = selectedRange.getRow();
  
  // Strategy 1: Try cell to the right of the selection
  let targetCol = lastCol + 1;
  let targetRow = firstRow;
  
  // Strategy 2: If we're dealing with a data table, place below the selection
  if (selectedRange.getNumRows() > 1 && selectedRange.getNumColumns() > 1) {
    targetRow = lastRow + 1;
    targetCol = selectedRange.getColumn();
  }
  
  // Strategy 3: For single row selections, place to the right
  if (selectedRange.getNumRows() === 1) {
    targetRow = firstRow;
    targetCol = lastCol + 1;
  }
  
  // Ensure we don't go beyond reasonable bounds
  const maxCols = sheet.getMaxColumns();
  const maxRows = sheet.getMaxRows();
  
  if (targetCol > maxCols) {
    targetCol = maxCols;
  }
  if (targetRow > maxRows) {
    targetRow = maxRows;
  }
  
  return sheet.getRange(targetRow, targetCol).getA1Notation();
}

// Helper function for testing and debugging
function previewFormulaGeneration(promptText) {
  try {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
      return { error: "No range selected for context" };
    }
    
    const values = range.getValues();
    const analysis = analyzeSelectedData();
    const contextualPrompt = createContextualPrompt(promptText, analysis, values);
    const targetCell = findTargetCell(range);
    
    return {
      userPrompt: promptText,
      dataAnalysis: analysis,
      contextualPrompt: contextualPrompt,
      targetCell: targetCell,
      sampleData: values.slice(0, 3)
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

// Helper function for data type detection
function detectDataType(values) {
  if (!values || values.length === 0) return 'empty';
  
  const numberCount = values.filter(val => typeof val === 'number').length;
  const textCount = values.filter(val => typeof val === 'string' && val !== "").length;
  const dateCount = values.filter(val => val instanceof Date).length;
  
  const total = values.length;
  
  if (dateCount / total > 0.5) return 'date';
  if (numberCount / total > 0.7) return 'number';
  if (textCount / total > 0.7) return 'text';
  return 'mixed';
}

