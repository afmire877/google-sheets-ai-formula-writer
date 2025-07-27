const tools = [
  {
    name: "analyzeDataStructure",
    description: "Analyzes the selected data to understand column types, headers, and data patterns",
    parameters: {
      type: "object",
      properties: {},
      required: []
    },
    function: () => {
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
    }
  },

  {
    name: "getSelectedRangeDetails",
    description: "Gets detailed information about the currently selected range including headers and data context",
    parameters: {
      type: "object",
      properties: {},
      required: []
    },
    function: () => {
      const range = SpreadsheetApp.getActiveRange();
      if (!range) return { error: "No range selected" };

      const values = range.getValues();
      const sheet = range.getSheet();

      return {
        sheetName: sheet.getName(),
        rangeAddress: range.getA1Notation(),
        numRows: values.length,
        numColumns: values[0].length,
        headers: values[0],
        sampleData: values.slice(1, 6),
        allData: values,
        isEmpty: values.every(row => row.every(cell => cell === "")),
        containsNumbers: values.flat().some(val => typeof val === 'number'),
        containsText: values.flat().some(val => typeof val === 'string' && val !== ""),
        containsDates: values.flat().some(val => val instanceof Date)
      };
    }
  },

  {
    name: "generateFormula",
    description: "Generates a Google Sheets formula based on user request and data context",
    parameters: {
      type: "object",
      properties: {
        request: {
          type: "string",
          description: "The user's natural language request for what formula they need"
        },
        targetCell: {
          type: "string",
          description: "Optional: specific cell where formula should be placed (e.g., 'D1')"
        }
      },
      required: ["request"]
    },
    function: (args) => {
      const range = SpreadsheetApp.getActiveRange();
      if (!range) return { error: "No range selected for context" };

      const values = range.getValues();
      const rangeAddress = range.getA1Notation();

      // This is a placeholder - in real implementation, this would use AI to generate formula
      // For now, return the context that would be sent to AI
      return {
        userRequest: args.request,
        dataContext: {
          rangeAddress: rangeAddress,
          headers: values[0],
          sampleData: values.slice(1, 3),
          dataTypes: analyzeColumnTypes(values)
        },
        suggestedFormula: "=PLACEHOLDER_FORMULA",
        explanation: "Formula generation logic would be implemented here",
        targetCell: args.targetCell
      };
    }
  },

  {
    name: "insertFormula",
    description: "Inserts a formula into a specific cell",
    parameters: {
      type: "object",
      properties: {
        formula: {
          type: "string",
          description: "The formula to insert (including = sign)"
        },
        cellAddress: {
          type: "string",
          description: "Cell address where to insert formula (e.g., 'D1')"
        }
      },
      required: ["formula", "cellAddress"]
    },
    function: (args) => {
      try {
        const sheet = SpreadsheetApp.getActiveSheet();
        const cell = sheet.getRange(args.cellAddress);
        cell.setFormula(args.formula);

        return {
          success: true,
          message: `Formula ${args.formula} inserted into cell ${args.cellAddress}`,
          cellValue: cell.getValue()
        };
      } catch (error) {
        return {
          success: false,
          error: error.toString(),
          message: "Failed to insert formula"
        };
      }
    }
  },

  {
    name: "getColumnInfo",
    description: "Gets detailed information about a specific column in the selected range",
    parameters: {
      type: "object",
      properties: {
        columnIndex: {
          type: "number",
          description: "Zero-based index of the column to analyze"
        }
      },
      required: ["columnIndex"]
    },
    function: (args) => {
      const range = SpreadsheetApp.getActiveRange();
      if (!range) return { error: "No range selected" };

      const values = range.getValues();
      const colIndex = args.columnIndex;

      if (colIndex >= values[0].length) {
        return { error: "Column index out of range" };
      }

      const columnData = values.map(row => row[colIndex]);
      const dataValues = columnData.slice(1).filter(val => val !== "");

      return {
        columnIndex: colIndex,
        header: columnData[0],
        totalValues: dataValues.length,
        uniqueValues: [...new Set(dataValues)],
        dataType: detectDataType(dataValues),
        min: Math.min(...dataValues.filter(val => typeof val === 'number')),
        max: Math.max(...dataValues.filter(val => typeof val === 'number')),
        sum: dataValues.filter(val => typeof val === 'number').reduce((a, b) => a + b, 0),
        average: dataValues.filter(val => typeof val === 'number').reduce((a, b) => a + b, 0) / dataValues.filter(val => typeof val === 'number').length
      };
    }
  },

  {
    name: "suggestFormulaType",
    description: "Suggests the most appropriate formula type based on user request",
    parameters: {
      type: "object",
      properties: {
        request: {
          type: "string",
          description: "User's natural language request"
        }
      },
      required: ["request"]
    },
    function: (args) => {
      const request = args.request.toLowerCase();
      const suggestions = [];

      if (request.includes('sum') || request.includes('total')) {
        if (request.includes('if') || request.includes('where') || request.includes('condition')) {
          suggestions.push({ type: 'SUMIF/SUMIFS', confidence: 'high', reason: 'Conditional sum requested' });
        } else {
          suggestions.push({ type: 'SUM', confidence: 'high', reason: 'Simple sum requested' });
        }
      }

      if (request.includes('lookup') || request.includes('find') || request.includes('match')) {
        suggestions.push({ type: 'VLOOKUP/XLOOKUP', confidence: 'high', reason: 'Lookup operation requested' });
      }

      if (request.includes('count')) {
        if (request.includes('if') || request.includes('where')) {
          suggestions.push({ type: 'COUNTIF/COUNTIFS', confidence: 'high', reason: 'Conditional count requested' });
        } else {
          suggestions.push({ type: 'COUNT/COUNTA', confidence: 'medium', reason: 'Count operation requested' });
        }
      }

      if (request.includes('average') || request.includes('mean')) {
        suggestions.push({ type: 'AVERAGE/AVERAGEIF', confidence: 'high', reason: 'Average calculation requested' });
      }

      if (request.includes('rank') || request.includes('order')) {
        suggestions.push({ type: 'RANK', confidence: 'medium', reason: 'Ranking requested' });
      }

      return {
        userRequest: args.request,
        suggestions: suggestions,
        topSuggestion: suggestions.length > 0 ? suggestions[0] : { type: 'CUSTOM', confidence: 'low', reason: 'Custom formula needed' }
      };
    }
  }
];

// Helper functions
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

function analyzeColumnTypes(values) {
  if (!values || values.length === 0) return [];

  return values[0].map((_, colIndex) => {
    const columnData = values.slice(1).map(row => row[colIndex]);
    return detectDataType(columnData);
  });
}
