# Google Sheets AI Formula Writer

A Google Apps Script project that uses AI to automatically generate Google Sheets formulas based on natural language descriptions.

## Project Overview

This project provides AI-powered formula generation for Google Sheets. Simply describe what you want to calculate in plain English, and the AI will write the appropriate Google Sheets formula for you. It integrates with OpenAI's API to understand your data context and generate accurate formulas for any calculation need.

## Features

- **Natural Language Formula Generation**: Convert plain English descriptions into Google Sheets formulas
- **Context-Aware Calculations**: AI understands your selected data and generates appropriate formulas
- **Complex Formula Support**: Handles SUMIF, VLOOKUP, QUERY, pivot calculations, and more
- **Custom Menu Integration**: Easy-to-use interface directly in Google Sheets
- **Smart Data Analysis**: AI can suggest formulas based on your data structure

## File Structure

```
/
├── Code.js              # Main script with OpenAI integration and dashboard functions
├── helpers.js           # Utility functions for data selection and processing
├── tools.js            # AI tool definitions for function calling
├── promptmodal.html    # User interface for AI prompts
├── appsscript.json     # Project configuration
└── CLAUDE.md          # This documentation file
```

### File Descriptions

- **Code.js**: Contains the core functionality including OpenAI API integration and formula generation functions
- **helpers.js**: Helper functions for working with selected ranges, data analysis, and the custom menu system
- **tools.js**: Defines AI tools that can be called by OpenAI for understanding data context and structure
- **promptmodal.html**: Simple HTML interface for describing what formula you need

## Setup Instructions

### Prerequisites
- Google Apps Script project
- OpenAI API key

### Configuration
1. **API Key Setup** (Security Critical):
   ```javascript
   // Instead of hardcoding in Code.js, use PropertiesService:
   PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'your-key-here');
   ```

2. **Enable Required Services**:
   - Google Sheets API (automatically enabled)
   - URL Fetch service (for OpenAI API calls)

3. **Deploy as Web App** (if needed):
   - Set execute permissions and access controls in appsscript.json

## Usage Examples

### Common Formula Requests

1. **Sum and Count Formulas**:
   - "Sum all values in column C where column B equals 'Groceries'"
   - "Count how many rows have a value greater than 100"
   - "Calculate the average of the selected data"

2. **Lookup and Match Formulas**:
   - "Find the price for item X in this table"
   - "Get the corresponding value from another column"
   - "Create a formula to lookup data from another sheet"

3. **Conditional and Complex Formulas**:
   - "Calculate total only if date is in current month"
   - "Sum values based on multiple criteria"
   - "Create a formula that ranks these values from highest to lowest"

4. **Data Analysis Formulas**:
   - "Show me the top 5 highest values"
   - "Calculate percentage of each item compared to total"
   - "Find duplicates in this data range"

### How It Works

1. **Select Your Data**: Highlight the range of cells you want to work with
2. **Open AI Menu**: Use the "My Helpers" → "Create with AI" menu option
3. **Describe Your Need**: Type what calculation you want in plain English
4. **Get Your Formula**: The AI generates the appropriate Google Sheets formula
5. **Apply the Formula**: The formula is automatically inserted into your sheet

The system works with any data format - just describe what you want to calculate!

## Development Guidelines

### Adding New AI Tools

1. **Define the tool in tools.js**:
   ```javascript
   {
     name: "newToolName",
     description: "What this tool does",
     parameters: { /* parameter schema */ },
     function: (args) => { /* implementation */ }
   }
   ```

2. **Use in prompts**: The AI will automatically call available tools when appropriate

### Security Best Practices

- **Never commit API keys**: Use PropertiesService for sensitive data
- **Validate user inputs**: Sanitize prompts before sending to AI
- **Limit access**: Configure appropriate sharing permissions
- **Monitor usage**: Track API calls to prevent abuse

### Error Handling

The system includes basic error handling for:
- Missing sheets or data
- API failures
- Invalid data formats

Expand error handling by wrapping functions in try-catch blocks and providing user-friendly error messages.

## Technical Details

### OpenAI Integration
- **Model**: Uses GPT-4 for complex formula generation, GPT-4o-mini for simple calculations
- **Function Calling**: Implements OpenAI's tool calling to understand data context
- **Token Management**: Limits responses to prevent excessive API usage

### Google Sheets Integration
- **Range Selection**: Works with user-selected data ranges to understand context
- **Formula Generation**: Creates accurate formulas based on data structure
- **Smart Insertion**: Places generated formulas in appropriate cells

## Troubleshooting

### Common Issues

1. **"No active range selected"**:
   - Ensure data is selected before running AI functions
   - Check that the sheet contains transaction data

2. **"AI response failed"**:
   - Verify OpenAI API key is correctly set
   - Check internet connection
   - Ensure API quota is not exceeded

3. **"Formula not working as expected"**:
   - Check that your data is properly selected
   - Verify column headers match your description
   - Try rephrasing your request more specifically

### Performance Tips

- Be specific in your formula requests for better results
- Select only the relevant data range before asking for formulas
- Use descriptive column headers to help AI understand your data structure

## Roadmap

### Planned Features
- Formula explanation and documentation
- Advanced formula suggestions based on data patterns
- Support for array formulas and advanced functions
- Integration with Google Sheets add-on marketplace
- Formula optimization suggestions

### Enhancement Ideas
- Natural language formula editing
- Formula template library
- Multi-sheet formula generation
- Custom function creation assistance

## Contributing

When making changes:
1. Test with various data types and structures
2. Verify generated formulas work correctly
3. Update this documentation for new features
4. Follow Google Apps Script best practices

## API Usage Notes

- OpenAI API calls are logged for debugging
- Consider implementing rate limiting for production use
- Monitor token usage to manage costs
- API responses are cached briefly to improve performance