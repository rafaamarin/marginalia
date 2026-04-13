import sys
import argparse
import os
import warnings
from pathlib import Path

# Suppress noisy ebooklib warnings that can interfere with the TUI
warnings.filterwarnings('ignore', category=UserWarning, module='ebooklib')

from pypdf import PdfReader
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, TextArea, Label
from textual.containers import Horizontal, VerticalScroll
from textual.binding import Binding

class Marginalia(App):
    CSS = """
    Screen {
        background: #1a1b26;
    }
    #reader-container {
        width: 65%;
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
    Label {
        width: 100%;
    }
    """

    BINDINGS = [
        Binding("ctrl+q", "quit", "Quit", show=True),
        Binding("ctrl+s", "save_notes", "Save Notes", show=True),
    ]

    def __init__(self, content, source_name):
        super().__init__()
        self.content = content
        self.source_name = source_name
        self.notes_file = Path(f"{source_name}.notes.md")
        
    def compose(self) -> ComposeResult:
        yield Header()
        with Horizontal():
            with VerticalScroll(id="reader-container"):
                yield Label(self.content)
            with VerticalScroll(id="notes-container"):
                initial_notes = ""
                if self.notes_file.exists():
                    initial_notes = self.notes_file.read_text()
                yield TextArea(initial_notes, id="notes-area", language="markdown")
        yield Footer()

    def on_mount(self) -> None:
        self.title = f"Marginalia - {self.source_name}"
        self.sub_title = f"Saving to {self.notes_file.name}"
        self.query_one("#reader-container").focus()

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
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        return f"Error reading PDF: {e}"

def extract_epub(path):
    try:
        book = epub.read_epub(path)
        text = []
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text.append(soup.get_text())
        return "\n\n".join(text)
    except Exception as e:
        return f"Error reading EPUB: {e}"

def main():
    parser = argparse.ArgumentParser(description="Marginalia - A CLI reader with side-notes.")
    parser.add_argument("file", nargs="?", help="Path to .epub, .pdf, or text file. Use '-' or omit for stdin.")
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
        else:
            try:
                content = path.read_text()
            except UnicodeDecodeError:
                print("Binary file detected. Please provide a text, PDF, or EPUB file.")
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
