# Outlook AI Assistant

An Outlook add-in that allows you to process selected text with AI using custom instructions. Supports multiple LLM providers: OpenAI, Claude, Gemini, and custom OpenAI-compatible APIs.

## Features

- **Select & Process**: Select text in your email, process it with AI
- **Custom Instructions**: Enter any instruction for the AI
- **Saved Prompts**: Save frequently used prompts for quick access
- **Multi-Provider Support**: Works with OpenAI, Claude, Gemini, or custom APIs
- **Insert Results**: Automatically insert AI-generated text below your selection

## Installation

### Prerequisites

- Node.js (for running the local development server)
- Microsoft Outlook (Web, Windows, or Mac)

### Step 1: Start the Local Server

1. Open a terminal in this directory
2. Run:
   ```bash
   npm start
   ```
   This starts a local server at `https://localhost:3000`

### Step 2: Start the Local Server with HTTPS

For Outlook Desktop to notice the add-in, it **must** run over HTTPS with a trusted certificate.

1.  **Start the server with certificates:**
    ```bash
    npm run start:desktop
    ```
    This uses the Microsoft Office developer certificates we just installed.

2.  **Verify access:**
    Open [https://localhost:3000](https://localhost:3000) in your browser. If you don't see a "Your connection is not private" warning, you're good to go.

#### Alternative: Use ngrok (Recommended for external testing)
```bash
npx ngrok http 3000
```
Then update all URLs in `manifest.xml` to use the ngrok HTTPS URL.

### Step 3: Sideload the Add-in

#### Outlook on the Web
1. Go to [outlook.office.com](https://outlook.office.com)
2. Open any email or compose a new one
3. Click **...** (More actions) → **Get Add-ins**
4. Click **My add-ins** → **Add a custom add-in** → **Add from file**
5. Upload the `manifest.xml` file

#### Outlook Desktop (Windows)
1. Open Outlook
2. Go to **File** → **Manage Add-ins** (opens in browser)
3. Click **My add-ins** → **Add a custom add-in** → **Add from file**
4. Upload the `manifest.xml` file

## Usage

1. **Compose an email** in Outlook
2. **Select some text** you want to process
3. Click the **AI Assistant** button in the ribbon
4. In the task pane:
   - Your selected text appears at the top
   - Choose a saved prompt OR enter a custom instruction
   - Click **Process with AI**
5. Review the result and click **Insert Below Selection**

## Configuration

Click the **⚙️ Settings** button to:

### LLM Provider Setup
1. Select your provider (OpenAI, Claude, Gemini, or Custom)
2. Enter your API key
3. Optionally modify the endpoint/model
4. Click **Save Settings**

### Saved Prompts
- Click **Add** to create a new saved prompt
- Click the edit icon to modify existing prompts
- Click the delete icon to remove prompts

## File Structure

```
├── manifest.xml           # Outlook add-in manifest
├── package.json           # Project config
├── README.md              # This file
├── assets/                # Icons
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── src/
    ├── lib/
    │   ├── llm-providers.js   # LLM API integrations
    │   ├── prompt-manager.js  # Saved prompts CRUD
    │   └── storage.js         # localStorage wrapper
    ├── taskpane/
    │   ├── taskpane.html      # Main UI
    │   ├── taskpane.css       # Styles
    │   └── taskpane.js        # Logic
    └── settings/
        ├── settings.html      # Settings UI
        ├── settings.css       # Styles
        └── settings.js        # Logic
```

## Troubleshooting

### "Could not get selection"
- Make sure you're in compose mode (writing a new email or reply)
- The add-in doesn't work in read mode

### "No API key configured"
- Go to Settings and configure your LLM provider with a valid API key

### Add-in doesn't appear
- Make sure the local server is running (`npm start`)
- Check that you're using HTTPS
- Try refreshing Outlook

## License

MIT
