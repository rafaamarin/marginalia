# Marginalia 📚

A minimalist CLI ebook reader for Linux.

### Features
- Supports .epub, .pdf, .docx, and raw text.
- Parses and renders tables, links, and formatting as Markdown.
- Split-screen: Read on the left, type Markdown notes on the right.
- Auto-saves progress and reading position (`~/.marginalia/state.json`).
- Interactive Table of Contents sidebar (`Ctrl+T`).
- In-document search (`Ctrl+F`).
- Built-in Dictionary lookup (`Ctrl+D`).

### Install
1. Clone this repo.
2. Set up a virtual environment and install dependencies: `pip install textual pypdf ebooklib beautifulsoup4 rich pymupdf4llm python-docx markdownify requests`
3. Run `python marginalia.py <file_path>`
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/d50fd7af-8d44-46eb-85de-b77f62344b50" />
