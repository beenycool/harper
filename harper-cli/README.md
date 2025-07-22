# `harper-cli`

## What?

`harper-cli` is a small, experimental frontend for Harper.
It can be used in any situation where you might need to check a large number of files automatically (like in continuous integration).

## Features

### Linting
- Lint documents for grammar and style issues
- Support for multiple file formats (Markdown, HTML, etc.)
- Configurable rules and dialect support

### Word Management
- **Add words to user dictionary**: `harper-cli add-word "customword"`
- **Remove words from user dictionary**: `harper-cli remove-word "customword"`
- **List user dictionary words**: `harper-cli list-user-words`
- Support for custom dictionary paths
- Case-insensitive word management
- Automatic validation of word characters

### Dictionary Tools
- List all words in Harper's dictionary: `harper-cli words`
- Get metadata for specific words: `harper-cli metadata "word"`
- Analyze word forms and affixes: `harper-cli forms "word/annotation"`

### Analysis Tools
- Parse documents and show token information
- Mine words from documents by frequency
- Extract compound words and nominal phrases

## Usage Examples

```bash
# Check a document for issues
harper-cli lint document.md

# Add a technical term to your personal dictionary
harper-cli add-word "TypeScript"

# Remove a word you no longer want in your dictionary
harper-cli remove-word "oldterm"

# List all words in your personal dictionary
harper-cli list-user-words

# Use custom dictionary path
harper-cli add-word "project-specific-term" --user-dict-path ./project-dict.txt
```

## Possible Future Features

- On-disk caching
- Machine-readable output
