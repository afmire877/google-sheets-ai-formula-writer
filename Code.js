// API key should be set in PropertiesService for security
// Use: PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'your-key-here');
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

function callOpenAIForFormula(promptText) {
  const apiKey = OPENAI_API_KEY;

  if (!apiKey) {
    throw new Error("OpenAI API key not found. Please set it using PropertiesService.");
  }

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: `You are a Google Sheets formula expert. Generate accurate Google Sheets formulas based on user requests and data context.

IMPORTANT RULES:
1. Always return the formula starting with = sign
2. Use Google Sheets function names (not Excel)
3. Reference ranges using A1 notation 
4. For relative references, use the range structure shown in the data context
5. Be precise with column references based on headers provided
6. Return ONLY the formula, no explanations unless specifically asked

Common Google Sheets functions:
- SUM, SUMIF, SUMIFS for totaling
- COUNT, COUNTA, COUNTIF, COUNTIFS for counting  
- AVERAGE, AVERAGEIF, AVERAGEIFS for averages
- VLOOKUP, XLOOKUP, INDEX/MATCH for lookups
- QUERY for complex filtering and analysis
- FILTER for dynamic filtering`
      },
      { role: "user", content: promptText }
    ],
    temperature: 0.1,
    max_tokens: 500
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`OpenAI API error: ${result.error?.message || 'Unknown error'}`);
  }

  return result.choices[0]?.message?.content?.trim() || null;
}

function showPromptModal() {
  const html = HtmlService.createHtmlOutputFromFile('promptmodal')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'AI Formula Assistant');
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sheet copilot')
    .addItem('Log Selected Range', 'logSelectedRange')
    .addItem('Write Summary Row', 'writeSummaryRow')
    .addSeparator()
    .addItem('🧠 Create with AI', 'showPromptModal')
    .addSeparator()
    .addItem('🔍 Debug: Preview Context', 'debugPreviewContext')
    .addToUi();
}

// Add-on specific functions
function onInstall(e) {
  onOpen(e);
}

function onHomepage(e) {
  return createCard('Select data and describe the formula you need');
}

function onSelectionChange(e) {
  if (!e || !e.commonEventObject) return;
  
  try {
    const range = SpreadsheetApp.getActiveRange();
    if (range && range.getNumRows() > 0 && range.getNumColumns() > 0) {
      const analysis = analyzeSelectedData();
      const message = `Selected: ${analysis.rangeAddress} (${analysis.rowCount}×${analysis.colCount})`;
      return createCard(message, true);
    }
  } catch (error) {
    return createCard('Select data to get started');
  }
  
  return createCard('Select data to get started');
}

function createCard(message, hasData = false) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('AI Formula Writer')
      .setSubtitle('Generate formulas from natural language'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(message)));

  if (hasData) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('🧠 Create Formula')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showPromptModal')))));
  }

  return card.build();
}

function logSelectedRange() {
  const values = getSelectedValues();
  Logger.log(values);
}

function debugPreviewContext() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range of data first');
    return;
  }

  const analysis = analyzeSelectedData();
  const preview = previewFormulaGeneration("Sum all values in the Amount column");

  let message = `DEBUG INFO:\n\n`;
  message += `Selected Range: ${analysis.rangeAddress}\n`;
  message += `Rows: ${analysis.rowCount}, Columns: ${analysis.colCount}\n`;
  message += `Has Headers: ${analysis.hasHeaders}\n\n`;

  if (analysis.columns) {
    message += `Columns:\n`;
    analysis.columns.forEach((col, i) => {
      message += `  ${String.fromCharCode(65 + i)}: "${col.header}" (${col.dataType})\n`;
    });
  }

  message += `\nTarget Cell: ${preview.targetCell}\n`;
  message += `\nThis info will be sent to AI when generating formulas.`;

  SpreadsheetApp.getUi().alert('Data Context Preview', message, SpreadsheetApp.getUi().ButtonSet.OK);
}


