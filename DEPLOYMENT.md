# Google Workspace Add-on Deployment Guide

## For Public Distribution

### 1. Prepare for Google Workspace Marketplace

**Requirements:**
- Google Cloud Console project
- OAuth consent screen configured
- Add-on reviewed by Google (for public listing)

**Steps:**
1. **Create Google Cloud Project**:
   ```bash
   # Link your Apps Script to Google Cloud Console
   clasp setting scriptId YOUR_SCRIPT_ID
   ```

2. **Configure OAuth Consent Screen**:
   - Go to Google Cloud Console → APIs & Services → OAuth consent screen
   - Fill out app information, privacy policy, terms of service
   - Add required scopes:
     - `https://www.googleapis.com/auth/spreadsheets.currentonly`
     - `https://www.googleapis.com/auth/script.external_request`

3. **Deploy as Add-on**:
   ```bash
   clasp deploy --description "AI Formula Writer Add-on"
   ```

### 2. Private Distribution (Easier)

**For sharing with specific users/organizations:**

1. **Deploy as Web App**:
   ```bash
   clasp deploy --description "AI Formula Writer" 
   ```

2. **Share deployment URL** with users

3. **Users install by**:
   - Opening the deployment URL
   - Granting permissions
   - Add-on appears in their Google Sheets

### 3. Domain-wide Installation (Google Workspace)

**For Google Workspace admins:**

1. **Admin Console** → Apps → Google Workspace Marketplace apps
2. **Add app** → Enter your deployment URL
3. **Configure** permissions and user access
4. **Install** for all users in domain

## Installation Methods for Users

### Method 1: Direct Installation Link
Share this URL format with users:
```
https://script.google.com/macros/d/YOUR_SCRIPT_ID/edit?usp=sharing
```

### Method 2: Google Workspace Marketplace (after approval)
Users find and install from Google Workspace Marketplace

### Method 3: Manual Installation
Users copy the script code into their own Apps Script project

## Required User Setup

**Each user must:**
1. **Set OpenAI API Key**:
   ```javascript
   // Run once in Apps Script editor:
   PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'their-key');
   ```

2. **Grant permissions** when first running the add-on

## Security Considerations

- **API keys are private** - each user needs their own OpenAI key
- **No central API billing** - users pay for their own OpenAI usage
- **Minimal permissions** - only accesses current spreadsheet
- **No data storage** - formulas generated in real-time

## Publishing to Google Workspace Marketplace

**Requirements for public listing:**
- Privacy policy and terms of service
- App verification by Google
- Branding assets (logo, screenshots)
- Detailed app description
- Security assessment

**Timeline:**
- Initial review: 3-6 weeks
- Updates: 1-2 weeks

The fastest way to share with others is **Method 2 (Private Distribution)** - deploy and share the URL directly.