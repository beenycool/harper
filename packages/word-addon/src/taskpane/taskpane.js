/* Office.js API initialization */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("check-document").onclick = checkDocument;
    document.getElementById("check-selection").onclick = checkSelection;
    document.getElementById("add-word").onclick = addWord;
    
    // Initialize user dictionary display
    loadUserWords();
    
    // Allow adding words by pressing Enter
    document.getElementById("new-word").addEventListener("keypress", function(event) {
      if (event.key === "Enter") {
        addWord();
      }
    });
  }
});

// Import Harper.js functionality
// For production, Harper.js would be bundled or loaded from CDN
// This is a placeholder for the actual Harper integration
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

// User dictionary management
let userWords = [];

function addWord() {
  const wordInput = document.getElementById("new-word");
  const word = wordInput.value.trim();
  
  if (!word) {
    showStatus('Please enter a word to add', 'error');
    return;
  }
  
  // Validate word format (letters, numbers, apostrophes, hyphens)
  const wordPattern = /^[a-zA-Z0-9'-]+$/;
  if (!wordPattern.test(word)) {
    showStatus('Word can only contain letters, numbers, apostrophes, and hyphens', 'error');
    return;
  }
  
  // Check for duplicates (case-insensitive)
  const normalizedWord = word.toLowerCase();
  if (userWords.some(existingWord => existingWord.toLowerCase() === normalizedWord)) {
    showStatus(`'${word}' is already in your dictionary`, 'error');
    return;
  }
  
  // Add to user dictionary
  userWords.push(word);
  
  // Add to Harper linter if available
  if (harperLinter) {
    harperLinter.import_words([word]);
  }
  
  // Save to storage
  saveUserWords();
  
  // Update display
  loadUserWords();
  
  // Clear input
  wordInput.value = '';
  
  showStatus(`Added '${word}' to your dictionary`, 'success');
}

function removeWord(word) {
  userWords = userWords.filter(w => w !== word);
  saveUserWords();
  loadUserWords();
  
  // Note: Harper.js doesn't have a remove_word method, so we'd need to
  // recreate the linter with the updated word list
  if (harperLinter) {
    harperLinter = new Harper.Linter(Harper.Dialect.American);
    harperLinter.import_words(userWords);
  }
  
  showStatus(`Removed '${word}' from your dictionary`, 'success');
}

function loadUserWords() {
  // Load from Office settings
  Office.context.document.settings.refreshAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const savedWords = Office.context.document.settings.get('harper-user-words');
      if (savedWords) {
        userWords = JSON.parse(savedWords);
      }
      
      // Update Harper linter
      if (harperLinter && userWords.length > 0) {
        harperLinter.import_words(userWords);
      }
      
      updateWordListDisplay();
    }
  });
}

function saveUserWords() {
  Office.context.document.settings.set('harper-user-words', JSON.stringify(userWords));
  Office.context.document.settings.saveAsync();
}

function updateWordListDisplay() {
  const wordList = document.getElementById("word-list");
  
  if (userWords.length === 0) {
    wordList.innerHTML = '<p class="instruction">No words in your dictionary yet.</p>';
    return;
  }
  
  wordList.innerHTML = userWords.map(word => `
    <div class="word-item">
      <span class="word-text">${word}</span>
      <button class="remove-word" onclick="removeWord('${word}')">Remove</button>
    </div>
  `).join('');
}

function checkDocument() {
  if (!harperLinter) {
    showStatus('Harper grammar checker is not ready', 'error');
    return;
  }
  
  showStatus('Checking document...', 'info');
  
  Word.run(async (context) => {
    // Get the entire document text
    const body = context.document.body;
    context.load(body, 'text');
    
    await context.sync();
    
    const text = body.text;
    const lints = harperLinter.lint(text, harperLinter.Language.Plain);
    
    displayLintResults(lints, text);
    
    if (lints.length === 0) {
      showStatus('No grammar issues found!', 'success');
    } else {
      showStatus(`Found ${lints.length} grammar issue(s)`, 'info');
    }
  }).catch(error => {
    console.error('Error checking document:', error);
    showStatus('Error checking document', 'error');
  });
}

function checkSelection() {
  if (!harperLinter) {
    showStatus('Harper grammar checker is not ready', 'error');
    return;
  }
  
  showStatus('Checking selection...', 'info');
  
  Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    context.load(selection, 'text');
    
    await context.sync();
    
    if (!selection.text) {
      showStatus('No text selected', 'error');
      return;
    }
    
    const text = selection.text;
    const lints = harperLinter.lint(text, harperLinter.Language.Plain);
    
    displayLintResults(lints, text);
    
    if (lints.length === 0) {
      showStatus('No grammar issues found in selection!', 'success');
    } else {
      showStatus(`Found ${lints.length} grammar issue(s) in selection`, 'info');
    }
  }).catch(error => {
    console.error('Error checking selection:', error);
    showStatus('Error checking selection', 'error');
  });
}

function displayLintResults(lints, originalText) {
  const resultsSection = document.getElementById("results-section");
  const lintResults = document.getElementById("lint-results");
  
  if (lints.length === 0) {
    resultsSection.style.display = 'none';
    return;
  }
  
  resultsSection.style.display = 'block';
  
  lintResults.innerHTML = lints.map((lint, index) => {
    const span = lint.span();
    const problemText = originalText.substring(span.start, span.end);
    const suggestions = lint.suggestions();
    
    return `
      <div class="lint-item">
        <div class="lint-message">${lint.message()}</div>
        <div class="lint-problem">"${problemText}"</div>
        ${suggestions.length > 0 ? `
          <div class="lint-suggestions">
            ${suggestions.map(suggestion => `
              <button class="suggestion-button" onclick="applySuggestion(${index}, '${suggestion.get_replacement_text()}', ${span.start}, ${span.end})">
                ${suggestion.get_replacement_text() || '(remove)'}
              </button>
            `).join('')}
          </div>
        ` : ''}
      </div>
    `;
  }).join('');
}

function applySuggestion(lintIndex, replacementText, start, end) {
  Word.run(async (context) => {
    // Get the document body
    const body = context.document.body;
    
    // Create a range for the problematic text
    const range = body.getRange().getRange(start, end);
    
    // Replace the text
    range.insertText(replacementText, Word.InsertLocation.replace);
    
    await context.sync();
    
    showStatus('Suggestion applied', 'success');
    
    // Re-check the document after applying suggestion
    setTimeout(checkDocument, 500);
  }).catch(error => {
    console.error('Error applying suggestion:', error);
    showStatus('Error applying suggestion', 'error');
  });
}

function showStatus(message, type = 'info') {
  const status = document.getElementById("status");
  status.textContent = message;
  status.className = `status-message status-${type}`;
  
  // Clear status after 5 seconds
  setTimeout(() => {
    status.textContent = '';
    status.className = 'status-message';
  }, 5000);
}

// Make functions globally available for onclick handlers
window.removeWord = removeWord;
window.applySuggestion = applySuggestion;