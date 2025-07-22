/* Office.js API initialization for command functions */
Office.onReady(() => {
  // No UI initialization needed for command functions
});

// Import Harper.js functionality
// For production, Harper.js would be bundled or loaded from CDN
let harperLinter;

// Placeholder Harper implementation for demonstration
class HarperPlaceholder {
  constructor() {
    this.Language = { Plain: 'plain' };
    this.userWords = [];
  }
  
  lint(text, language) {
    // This is a simplified example - in production, this would use actual Harper.js
    const words = text.split(/\s+/);
    const issues = [];
    
    // Example: detect repeated words
    for (let i = 0; i < words.length - 1; i++) {
      if (words[i].toLowerCase() === words[i + 1].toLowerCase()) {
        issues.push({
          message: () => `Repeated word: "${words[i]}"`,
          span: () => ({ start: i * 10, end: (i + 1) * 10 }), // Simplified span calculation
          suggestions: () => [{ get_replacement_text: () => words[i] }]
        });
      }
    }
    
    return issues;
  }
  
  import_words(words) {
    this.userWords = [...this.userWords, ...words];
  }
}

// Initialize placeholder Harper
harperLinter = new HarperPlaceholder();

function loadUserWordsForCommands() {
  Office.context.document.settings.refreshAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const savedWords = Office.context.document.settings.get('harper-user-words');
      if (savedWords && harperLinter) {
        const userWords = JSON.parse(savedWords);
        harperLinter.import_words(userWords);
      }
    }
  });
}

/**
 * Quick grammar check command - this is called when the "Check Grammar" button is clicked
 * This provides a quick check without showing the full task pane
 */
function checkGrammar(event) {
  if (!harperLinter) {
    showNotification('Harper grammar checker is not ready');
    event.completed();
    return;
  }
  
  Word.run(async (context) => {
    try {
      // Get the current selection or entire document
      let textToCheck;
      let range;
      
      const selection = context.document.getSelection();
      context.load(selection, 'text');
      await context.sync();
      
      if (selection.text && selection.text.trim()) {
        // Check selection if text is selected
        textToCheck = selection.text;
        range = selection;
      } else {
        // Check entire document if no selection
        const body = context.document.body;
        context.load(body, 'text');
        await context.sync();
        textToCheck = body.text;
        range = body;
      }
      
      // Run Harper linting
      const lints = harperLinter.lint(textToCheck, harperLinter.Language.Plain);
      
      if (lints.length === 0) {
        showNotification('No grammar issues found!');
      } else {
        // Highlight first issue and show notification
        const firstLint = lints[0];
        const span = firstLint.span();
        
        // Create a range for the first problematic text
        const problemRange = range.getRange().getRange(span.start, span.end);
        
        // Highlight the problematic text
        problemRange.font.highlightColor = '#FFFF00'; // Yellow highlight
        
        await context.sync();
        
        showNotification(`Found ${lints.length} grammar issue(s). First issue: ${firstLint.message()}`);
      }
      
      event.completed();
    } catch (error) {
      console.error('Error in quick grammar check:', error);
      showNotification('Error checking grammar');
      event.completed();
    }
  });
}

function showNotification(message) {
  Office.context.ui.displayDialogAsync(
    `data:text/html,<html><body style="font-family: 'Segoe UI', sans-serif; padding: 20px; text-align: center;">
      <h3>Harper Grammar Checker</h3>
      <p>${message}</p>
      <button onclick="window.close()" style="background: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;">OK</button>
    </body></html>`,
    { height: 200, width: 400 }
  );
}

// Register the function with Office.js
Office.actions.associate("checkGrammar", checkGrammar);