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
    
    // Call OpenAI with the enhanced prompt and tools
    const result = callOpenAIWithFunction(contextualPrompt);
    
    if (result && result.choices && result.choices[0]) {
      const response = result.choices[0].message.content;
      
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
    } else {
      return {
        success: false,
        error: "No response from AI"
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
  let prompt = `You are a Google Sheets formula expert. Generate a formula based on this request:

`;
  prompt += `USER REQUEST: ${userRequest}\n\n`;
  prompt += `DATA CONTEXT:\n`;
  prompt += `- Range: ${analysis.rangeAddress}\n`;
  prompt += `- Size: ${analysis.rowCount} rows × ${analysis.colCount} columns\n`;
  prompt += `- Headers: ${analysis.hasHeaders ? 'Yes' : 'No'}\n`;
  
  if (analysis.columns) {
    prompt += `- Columns:\n`;
    analysis.columns.forEach(col => {
      prompt += `  ${col.index + 1}. "${col.header}" (${col.dataType})\n`;
    });
  }
  
  prompt += `\nSAMPLE DATA (first 3 rows):\n`;
  values.slice(0, 3).forEach((row, i) => {
    prompt += `Row ${i + 1}: ${row.join(' | ')}\n`;
  });
  
  prompt += `\nPlease provide:\n`;
  prompt += `1. The exact Google Sheets formula (starting with =)\n`;
  prompt += `2. A brief explanation of what it does\n`;
  prompt += `\nFormula:`;
  
  return prompt;
}

function extractFormulaFromResponse(response) {
  // Look for formula patterns in the response
  const formulaRegex = /=([^\n]+)/g;
  const matches = response.match(formulaRegex);
  
  if (matches && matches.length > 0) {
    return matches[0].trim();
  }
  
  return null;
}

function findTargetCell(selectedRange) {
  // Find the next available cell to the right of the selection
  const sheet = selectedRange.getSheet();
  const lastCol = selectedRange.getLastColumn();
  const firstRow = selectedRange.getRow();
  
  // Use the cell immediately to the right of the selection
  const targetCol = lastCol + 1;
  
  return sheet.getRange(firstRow, targetCol).getA1Notation();
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

