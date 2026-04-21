import sys
import argparse
import os
import warnings
import json
import requests
from pathlib import Path

# Suppress noisy ebooklib warnings that can interfere with the TUI
warnings.filterwarnings('ignore', category=UserWarning, module='ebooklib')

import pymupdf4llm
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

import markdownify
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, TextArea, Label, Markdown, Tree, Input
from textual.containers import Horizontal, VerticalScroll, Container
from textual.binding import Binding
from textual.screen import ModalScreen

# --- Modals ---

class DictionaryModal(ModalScreen):
    """A modal to search for a word's definition."""

    CSS = """
    DictionaryModal {
        align: center middle;
    }
    #dict-container {
        width: 60%;
        height: auto;
        max-height: 70%;
        background: #24283b;
        border: solid $accent;
        padding: 1 2;
    }
    #dict-input {
        margin-bottom: 1;
    }
    """

    def compose(self) -> ComposeResult:
        with VerticalScroll(id="dict-container"):
            yield Label("Lookup Definition:")
            yield Input(placeholder="Enter word...", id="dict-input")
            yield Markdown("", id="dict-result")

    def on_mount(self) -> None:
        self.query_one("#dict-input").focus()

    async def on_input_submitted(self, event: Input.Submitted) -> None:
        word = event.value.strip()
        if not word:
            return
        
        result_md = self.query_one("#dict-result")
        result_md.update(f"Searching for **{word}**...")
        
        try:
            # Using Free Dictionary API
            resp = requests.get(f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}")
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, list) and len(data) > 0:
                    meanings = data[0].get("meanings", [])
                    out = f"### {word}\n"
                    for meaning in meanings[:2]:
                        pos = meaning.get("partOfSpeech", "")
                        out += f"\n**{pos}**\n"
                        defs = meaning.get("definitions", [])
                        for d in defs[:2]:
                            out += f"- {d.get('definition', '')}\n"
                    result_md.update(out)
                else:
                    result_md.update("No definition found.")
            else:
                result_md.update("Word not found.")
        except Exception as e:
            result_md.update(f"Error fetching definition: {e}")

class SearchModal(ModalScreen):
    """A modal to search text within the document."""

    CSS = """
    SearchModal {
        align: center middle;
    }
    #search-container {
        width: 50%;
        height: auto;
        background: #24283b;
        border: solid $accent;
        padding: 1 2;
    }
    """

    def compose(self) -> ComposeResult:
        with Container(id="search-container"):
            yield Label("Search Document:")
            yield Input(placeholder="Enter search term...", id="search-input")

    def on_mount(self) -> None:
        self.query_one("#search-input").focus()

    def on_input_submitted(self, event: Input.Submitted) -> None:
        term = event.value.strip()
        self.dismiss(term)

# --- App ---

class Marginalia(App):
    CSS = """
    Screen {
        background: #1a1b26;
    }
    #main-layout {
        height: 100%;
    }
    #toc-sidebar {
        width: 25%;
        border-right: solid $accent;
        display: none;
        height: 100%;
        background: #1a1b26;
    }
    #toc-sidebar.-visible {
        display: block;
    }
    #reader-container {
        width: 1fr;
        border-right: solid $accent;
        padding: 1 2;
    }
    #notes-container {
        width: 35%;
        padding: 1;
    }
    TextArea {
        height: 100%;
        border: none;
        background: #24283b;
    }
    Markdown {
        width: 100%;
    }
    Tree {
        height: 100%;
        background: #1a1b26;
    }
    """

    BINDINGS = [
        Binding("ctrl+q", "quit", "Quit", show=True),
        Binding("ctrl+s", "save_notes", "Save Notes", show=True),
        Binding("ctrl+t", "toggle_toc", "TOC", show=True),
        Binding("ctrl+f", "search", "Search", show=True),
        Binding("ctrl+d", "dictionary", "Dictionary", show=True),
    ]

    def __init__(self, content, source_name):
        super().__init__()
        self.content = content
        self.source_name = source_name
        self.state_file = Path.home() / ".marginalia" / "state.json"
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        
        notes_dir = Path.home() / "marginalia-notes"
        notes_dir.mkdir(parents=True, exist_ok=True)
        self.notes_file = notes_dir / f"{source_name}.notes.md"
        
    def compose(self) -> ComposeResult:
        yield Header()
        with Horizontal(id="main-layout"):
            with VerticalScroll(id="toc-sidebar"):
                yield Tree("Table of Contents")
            with VerticalScroll(id="reader-container"):
                yield Markdown(self.content, id="markdown-reader")
            with VerticalScroll(id="notes-container"):
                initial_notes = ""
                if self.notes_file.exists():
                    initial_notes = self.notes_file.read_text()
                yield TextArea(initial_notes, id="notes-area", language="markdown")
        yield Footer()

    async def on_mount(self) -> None:
        self.title = f"Marginalia - {self.source_name}"
        self.sub_title = f"Saving to {self.notes_file.name}"
        self.query_one("#reader-container").focus()

        # Build TOC
        tree = self.query_one(Tree)
        tree.root.expand()
        lines = self.content.split('\n')
        for line in lines:
            if line.startswith('#'):
                # Very basic TOC parsing
                title = line.lstrip('#').strip()
                tree.root.add_leaf(title, data=title)
        
        # Load State
        if self.state_file.exists():
            try:
                states = json.loads(self.state_file.read_text())
                if self.source_name in states:
                    scroll_y = states[self.source_name]
                    # We might need a slight delay for layout before scrolling
                    self.set_timer(0.1, lambda: self.query_one("#reader-container").scroll_to(y=scroll_y, animate=False))
            except Exception:
                pass

    def on_unmount(self) -> None:
        # Save State
        scroll_y = self.query_one("#reader-container").scroll_y
        states = {}
        if self.state_file.exists():
            try:
                states = json.loads(self.state_file.read_text())
            except Exception:
                pass
        states[self.source_name] = scroll_y
        self.state_file.write_text(json.dumps(states))

    def action_toggle_toc(self) -> None:
        sidebar = self.query_one("#toc-sidebar")
        if sidebar.has_class("-visible"):
            sidebar.remove_class("-visible")
        else:
            sidebar.add_class("-visible")
            self.query_one(Tree).focus()

    def action_search(self) -> None:
        def check_search(term: str):
            if term:
                lines = self.content.split('\n')
                for i, line in enumerate(lines):
                    if term.lower() in line.lower():
                        # Rough estimate for scrolling based on line position.
                        # It's an approximation since markdown height varies.
                        container = self.query_one("#reader-container")
                        container.scroll_to(y=i, animate=True)
                        self.notify(f"Scrolled to approximate match: '{term}'")
                        return
                self.notify("Term not found.")
                
        self.push_screen(SearchModal(), check_search)

    def action_dictionary(self) -> None:
        self.push_screen(DictionaryModal())

    def on_tree_node_selected(self, event: Tree.NodeSelected) -> None:
        # Simple TOC jump logic
        # Find the line that matches the TOC heading and scroll to it.
        if event.node.data:
            heading = event.node.data
            lines = self.content.split('\n')
            for i, line in enumerate(lines):
                if line.startswith('#') and line.lstrip('#').strip() == heading:
                    self.query_one("#reader-container").scroll_to(y=i, animate=True)
                    self.query_one("#reader-container").focus()
                    break

    def action_save_notes(self) -> None:
        notes = self.query_one("#notes-area").text
        self.notes_file.write_text(notes)
        self.notify(f"Notes saved to {self.notes_file.name}")

    def on_text_area_changed(self, event: TextArea.Changed) -> None:
        # Simple auto-save to ensure notes are never lost
        notes = event.text_area.text
        self.notes_file.write_text(notes)

def extract_pdf(path):
    try:
        # Use pymupdf4llm for high quality PDF to Markdown conversion
        md_text = pymupdf4llm.to_markdown(str(path))
        return md_text
    except Exception as e:
        return f"Error reading PDF: {e}"

def extract_epub(path):
    try:
        book = epub.read_epub(path)
        text = []
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                content = item.get_content().decode('utf-8')
                # Use markdownify to preserve formatting, tables, links
                md_text = markdownify.markdownify(content, heading_style="ATX")
                text.append(md_text)
        return "\\n\\n".join(text)
    except Exception as e:
        return f"Error reading EPUB: {e}"

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Unsupported parent type for iteration")
    
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_docx(path):
    try:
        doc = docx.Document(path)
        text_parts = []
        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                if block.text.strip():
                    text_parts.append(block.text)
            elif isinstance(block, Table):
                md_table = []
                for i, row in enumerate(block.rows):
                    row_data = [cell.text.replace('\\n', ' ').strip() for cell in row.cells]
                    md_table.append("| " + " | ".join(row_data) + " |")
                    if i == 0:
                        md_table.append("|" + "|".join(["---" for _ in row_data]) + "|")
                if md_table:
                    text_parts.append("\\n" + "\\n".join(md_table) + "\\n")
        return "\\n\\n".join(text_parts)
    except Exception as e:
        return f"Error reading DOCX: {e}"

def main():
    parser = argparse.ArgumentParser(description="Marginalia - A CLI reader with side-notes.")
    parser.add_argument("file", nargs="?", help="Path to .epub, .pdf, .docx, or text file. Use '-' or omit for stdin.")
    args = parser.parse_args()

    content = ""
    source_name = "session"

    if args.file and args.file != "-":
        path = Path(args.file)
        if not path.exists():
            print(f"File not found: {args.file}")
            sys.exit(1)
        
        source_name = path.stem
        suffix = path.suffix.lower()
        
        if suffix == ".pdf":
            content = extract_pdf(path)
        elif suffix == ".epub":
            content = extract_epub(path)
        elif suffix == ".docx":
            content = extract_docx(path)
        else:
            try:
                content = path.read_text()
            except UnicodeDecodeError:
                print("Binary file detected. Please provide a text, PDF, EPUB, or DOCX file.")
                sys.exit(1)
    else:
        # Read from stdin
        if not sys.stdin.isatty():
            content = sys.stdin.read()
            source_name = "stdin"
        else:
            parser.print_help()
            sys.exit(0)

    if not content.strip():
        print("No content to display.")
        sys.exit(1)

    app = Marginalia(content, source_name)
    app.run()

if __name__ == "__main__":
    main()