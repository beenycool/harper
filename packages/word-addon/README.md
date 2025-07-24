# Harper Word Add-in

A Microsoft Word Add-in that integrates Harper's grammar checking capabilities directly into Word documents.

## Features

- **Real-time Grammar Checking**: Check your entire document or selected text for grammar issues
- **Personal Dictionary Management**: Add and remove words from your personal dictionary within Word
- **Instant Feedback**: Get immediate suggestions for grammar improvements
- **Private and Secure**: All processing happens locally, your documents never leave your device
- **Seamless Integration**: Works directly within Microsoft Word's interface

## Installation

### Prerequisites

- Microsoft Word (Desktop or Web version)
- Node.js 14+ and npm/pnpm for development

### Development Setup

1. Clone the repository and navigate to the Word Add-in package:
   ```bash
   cd packages/word-addon
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm run dev-server
   ```

4. Sideload the Add-in in Word:
   ```bash
   npm run sideload
   ```

### Production Deployment

1. Build the Add-in:
   ```bash
   npm run build
   ```

2. Deploy the built files to your web server

3. Update the manifest.xml with your production URLs

4. Install the Add-in in Word using the updated manifest

## Usage

### Getting Started

1. Open Microsoft Word
2. Look for the "Harper" group in the Home tab ribbon
3. Click "Show Harper" to open the Harper task pane

### Checking Grammar

#### Full Document Check
1. Click "Check Document" in the Harper task pane
2. Review any grammar issues found in the results section
3. Click on suggested fixes to apply them automatically

#### Selection Check
1. Select the text you want to check
2. Click "Check Selection" in the Harper task pane
3. Review and apply suggestions as needed

#### Quick Check
- Use the "Check Grammar" button in the ribbon for a quick check
- This will highlight the first grammar issue found and show a notification

### Managing Your Personal Dictionary

#### Adding Words
1. Type a word in the "Add a word..." field
2. Click "Add" or press Enter
3. The word will be added to your personal dictionary and won't be flagged as an error

#### Removing Words
1. Find the word in your dictionary list
2. Click "Remove" next to the word
3. The word will be removed from your dictionary

### Word Validation

Words added to your dictionary must:
- Contain only letters, numbers, apostrophes, and hyphens
- Not be duplicates (case-insensitive)

## Technical Details

### Architecture

The Harper Word Add-in consists of:

- **Task Pane**: Main interface for grammar checking and dictionary management
- **Ribbon Commands**: Quick access buttons in Word's ribbon
- **Function Commands**: Background functions for quick grammar checks
- **Local Storage**: Personal dictionary stored in Word's document settings

### Integration with Harper.js

The Add-in is designed to use Harper.js (the WebAssembly version of Harper) to:
- Perform grammar checking locally in the browser
- Manage personal dictionaries
- Provide real-time suggestions

**Note**: The current implementation includes a placeholder Harper class for demonstration purposes. To use the full Harper functionality:

1. Build harper.js from the workspace: `cd ../harper.js && pnpm run build`
2. Update the taskpane.js to import the actual Harper.js library
3. Replace the HarperPlaceholder class with the real Harper implementation

Example integration:
```javascript
// Replace the placeholder with actual Harper.js import
import * as Harper from 'harper.js';
const harperLinter = new Harper.Linter(Harper.Dialect.American);
```

### Data Privacy

- All grammar checking happens locally in your browser
- Personal dictionary is stored in Word's local settings
- No data is sent to external servers
- Complete privacy and security of your documents

## Development

### Project Structure

```
src/
├── taskpane/          # Main task pane interface
│   ├── taskpane.html  # Task pane HTML
│   ├── taskpane.css   # Styling
│   └── taskpane.js    # Task pane functionality
├── commands/          # Ribbon command functions
│   ├── commands.html  # Commands page
│   └── commands.js    # Command implementations
assets/                # Icons and images
manifest.xml          # Add-in manifest
webpack.config.js     # Build configuration
```

### Available Scripts

- `npm run dev-server` - Start development server
- `npm run build` - Build for production
- `npm run sideload` - Sideload in Word for testing
- `npm run validate` - Validate manifest
- `npm run lint` - Run linting

### Customization

You can customize the Add-in by:
- Modifying the UI in `taskpane.html` and `taskpane.css`
- Adding new ribbon commands in `manifest.xml`
- Extending functionality in the JavaScript files
- Updating Harper configuration for different dialects

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Ensure the development server is running and Word trusts the certificate
2. **Harper not working**: Check browser console for errors related to Harper.js loading
3. **Dictionary not saving**: Verify Word has permission to save settings

### Debug Mode

Enable debug mode by:
1. Opening browser developer tools (F12)
2. Checking console for error messages
3. Using Word's Add-in debugging tools

## Contributing

Contributions are welcome! Please see the main Harper repository's contribution guidelines.

## License

This project is licensed under the Apache 2.0 License - see the main Harper repository for details.