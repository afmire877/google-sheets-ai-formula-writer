// API key should be set in PropertiesService for security
// Use: PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'your-key-here');
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

function callOpenAIWithFunction(promptText) {
  const apiKey = OPENAI_API_KEY;

  const payload = {
    model: "gpt-4-0613",
    messages: [
      { role: "user", content: promptText }
    ],
    tools: tools.map(tool => ({
      type: "function",
      function: {
        name: tool.name,
        description: tool.description,
        parameters: tool.parameters
      }
    })),
    tool_choice: "auto"
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

  const toolCall = result.choices[0]?.message?.tool_calls?.[0];
  if (toolCall && toolCall.function) {
    const { name, arguments: argsJSON } = toolCall.function;
    const fn = tools.find(t => t.name === name)?.function;
    const args = JSON.parse(argsJSON);
    if (fn) {
      const result = fn(args);
      Logger.log(`Function result for ${name}:`, result);
      return result;
    }
  }

  return result;
}

function showPromptModal() {
  const html = HtmlService.createHtmlOutputFromFile('promptmodal')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'AI Formula Assistant');
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('My Helpers')
    .addItem('Log Selected Range', 'logSelectedRange')
    .addItem('Write Summary Row', 'writeSummaryRow')
    .addItem('Create with AI', 'showPromptModal')
    .addToUi();
}

function logSelectedRange() {
  const values = getSelectedValues();
  Logger.log(values);
}


