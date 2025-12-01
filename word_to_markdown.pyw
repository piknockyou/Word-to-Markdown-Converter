#!/usr/bin/env python3
"""
Word to Markdown Converter
Uses MarkItDown library with automatic image extraction.
Supports drag & drop or GUI with optional page selection.
"""
import sys
import os
import subprocess
import tempfile
import re
import zipfile
import threading
import logging
from pathlib import Path
from typing import Any
from contextlib import contextmanager

# =================== LOGGING SETUP ===================
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

try:
    if sys.stdout and hasattr(sys.stdout, 'write'):
        _handler = logging.StreamHandler(sys.stdout)
        _handler.setFormatter(logging.Formatter('%(message)s'))
        logger.addHandler(_handler)
except Exception:
    pass

# =================== HIGH-DPI FIX (Windows) ===================
if sys.platform.startswith("win"):
    try:
        from ctypes import windll
        try:
            windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass
    except ImportError:
        pass

# =================== CONFIGURATION ===================
USE_FIXED_OUTPUT_FOLDER = False
FIXED_OUTPUT_FOLDER = r"C:\Markdown_Output"
MEDIA_FOLDER_NAME = "media"
AUTO_INSTALL_DEPS = True

# UI Configuration
FONT_MONO = ("Consolas", 11) if sys.platform.startswith("win") else ("Monaco", 11)
FONT_UI = ("Segoe UI", 10) if sys.platform.startswith("win") else ("Helvetica", 10)

# =================== AUTO-DEPENDENCY INSTALLATION ===================
def ensure_package(package_name: str, import_name: str | None = None) -> bool:
    """Ensure a package is installed, install if missing."""
    import_name = import_name or package_name.split('[')[0]
    try:
        __import__(import_name)
        return True
    except ImportError:
        try:
            logger.info(f"Installing {package_name}...")
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', package_name, '--quiet'],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            __import__(import_name)
            logger.info(f"Successfully installed {package_name}")
            return True
        except (subprocess.CalledProcessError, ImportError) as e:
            logger.warning(f"Could not install {package_name}: {e}")
            return False

def ensure_all_dependencies():
    """Install required dependencies."""
    ensure_package('markitdown', 'markitdown')
    if sys.platform.startswith('win'):
        ensure_package('pywin32', 'win32com')

# =================== CONSOLE MANAGEMENT ===================
@contextmanager
def managed_console():
    """Context manager to handle console allocation on Windows .pyw files."""
    if not sys.platform.startswith("win"):
        yield
        return
    
    import ctypes
    kernel32 = ctypes.windll.kernel32
    
    console_allocated = kernel32.AllocConsole()
    
    if not console_allocated:
        yield
        return
    
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    old_stdin = sys.stdin
    
    try:
        sys.stdout = open("CONOUT$", "w", encoding='utf-8')
        sys.stderr = open("CONOUT$", "w", encoding='utf-8')
        sys.stdin = open("CONIN$", "r", encoding='utf-8')
        
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(logging.Formatter('%(message)s'))
        logger.addHandler(handler)
        
        yield
    finally:
        for stream in [sys.stdout, sys.stderr, sys.stdin]:
            try:
                if stream and stream not in (old_stdout, old_stderr, old_stdin):
                    stream.close()
            except Exception:
                pass
        
        kernel32.FreeConsole()
        
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        sys.stdin = old_stdin

# =================== OPTIONAL IMPORTS ===================
# Check for pywin32 (Word COM) - Windows only
word_available = False
if sys.platform.startswith('win'):
    try:
        import win32com.client
        import pythoncom
        word_available = True
    except ImportError:
        pass

# Check for markitdown
markitdown_available = False
try:
    from markitdown import MarkItDown
    markitdown_available = True
except ImportError:
    pass

# =================== IMAGE EXTRACTION ===================
def extract_images_from_docx(docx_path: Path, output_dir: Path) -> list[Path]:
    """
    Extract all images from a .docx file to the output directory.
    Returns list of extracted image paths in order they appear in the docx.
    """
    extracted_images: list[Path] = []
    
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            # Get all media files, sorted by name to maintain order
            media_files = sorted([
                name for name in docx_zip.namelist()
                if name.startswith('word/media/')
            ])
            
            for item in media_files:
                image_name = Path(item).name
                image_data = docx_zip.read(item)
                
                output_path = output_dir / image_name
                with open(output_path, 'wb') as img_file:
                    img_file.write(image_data)
                
                extracted_images.append(output_path)
        
        if extracted_images:
            logger.info(f"Extracted {len(extracted_images)} images to {output_dir}")
    
    except Exception as e:
        logger.warning(f"Could not extract images from docx: {e}")
    
    return extracted_images

def fix_markdown_images(content: str, docx_path: Path, output_md_path: Path) -> str:
    """
    Fix broken image placeholders in markdown content.
    
    MarkItDown produces: ![](data:image/png;base64...)
    We replace these with: ![image1](media/image1.png)
    """
    # Pattern to match broken base64 placeholders
    # Matches: ![anything](data:image/type;base64...) where ... is literal dots or truncated
    broken_pattern = r'!\[([^\]]*)\]\(data:image/[^;]+;base64[^)]*\)'
    
    # Find all broken image references
    matches = list(re.finditer(broken_pattern, content))
    
    if not matches:
        logger.info("No broken image placeholders found")
        return content
    
    logger.info(f"Found {len(matches)} broken image placeholders to fix")
    
    # Extract actual images from docx
    media_dir = output_md_path.parent / MEDIA_FOLDER_NAME
    extracted_images = extract_images_from_docx(docx_path, media_dir)
    
    if not extracted_images:
        logger.warning("No images found in docx to extract")
        # Remove the broken placeholders since we can't fix them
        return re.sub(broken_pattern, '', content)
    
    # Replace placeholders with actual image links
    # Process in reverse order to maintain correct positions
    for i, match in enumerate(reversed(matches)):
        idx = len(matches) - 1 - i  # Original index
        
        if idx < len(extracted_images):
            image_path = extracted_images[idx]
            relative_path = f"{MEDIA_FOLDER_NAME}/{image_path.name}"
            alt_text = match.group(1) or f"image{idx + 1}"
            replacement = f"![{alt_text}]({relative_path})"
        else:
            # More placeholders than images - remove the extra ones
            replacement = ""
        
        content = content[:match.start()] + replacement + content[match.end():]
    
    return content

def remove_empty_folder(directory: Path) -> None:
    """Remove a folder if it's empty."""
    try:
        if directory.exists() and directory.is_dir():
            if not any(directory.iterdir()):
                directory.rmdir()
                logger.info(f"Removed empty folder: {directory}")
    except Exception as e:
        logger.warning(f"Could not remove folder {directory}: {e}")

# =================== WORD COM HELPER (Windows Only) ===================
class WordInstance:
    """Context manager for Microsoft Word COM automation (Windows only)."""
    
    def __init__(self):
        if not word_available:
            raise RuntimeError("Word automation requires pywin32 on Windows")
        self.app = None
    
    def __enter__(self):
        try:
            pythoncom.CoInitialize()
            self.app = win32com.client.Dispatch("Word.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            return self
        except Exception as e:
            pythoncom.CoUninitialize()
            raise RuntimeError(f"Failed to start Word: {e}")
    
    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any):
        if self.app:
            try:
                self.app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    
    def get_page_count(self, docx_path: Path) -> int:
        doc = None
        try:
            doc = self.app.Documents.Open(str(docx_path.absolute()))
            return doc.ComputeStatistics(2)
        finally:
            if doc:
                doc.Close(False)
    
    def extract_pages(self, input_path: Path, from_page: int, to_page: int) -> str:
        doc = None
        new_doc = None
        try:
            doc = self.app.Documents.Open(str(input_path.absolute()))
            
            temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
            os.close(temp_fd)
            
            start = doc.GoTo(1, 1, from_page)
            total_pages = doc.ComputeStatistics(2)
            
            if to_page < total_pages:
                end = doc.GoTo(1, 1, to_page + 1)
                rng = doc.Range(start.Start, end.Start)
            else:
                rng = doc.Range(start.Start, doc.Content.End)
            
            rng.Copy()
            
            new_doc = self.app.Documents.Add()
            new_doc.Range().Paste()
            new_doc.SaveAs(temp_path, 16)
            
            return temp_path
        finally:
            if new_doc:
                new_doc.Close(False)
            if doc:
                doc.Close(False)
    
    def convert_to_docx(self, doc_path: Path) -> str:
        doc = None
        try:
            doc = self.app.Documents.Open(str(doc_path.absolute()))
            temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
            os.close(temp_fd)
            doc.SaveAs(temp_path, 16)
            return temp_path
        finally:
            if doc:
                doc.Close(False)

def get_page_count_background(docx_path: Path) -> int | None:
    if not word_available:
        return None
    try:
        with WordInstance() as word:
            return word.get_page_count(docx_path)
    except Exception as e:
        logger.warning(f"Could not get page count: {e}")
        return None

# =================== CONVERSION ===================
def _get_output_path(input_path: Path) -> Path:
    if USE_FIXED_OUTPUT_FOLDER:
        output_dir = Path(os.path.expandvars(FIXED_OUTPUT_FOLDER))
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir / f"{input_path.stem}.md"
    return input_path.with_suffix('.md')

def _validate_input_file(input_path: Path) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"File not found: {input_path}")
    if input_path.suffix.lower() not in ['.doc', '.docx']:
        raise ValueError(f"File must be .doc or .docx, got: {input_path.suffix}")

def convert_file(
    input_path: str | Path,
    from_page: int | None = None,
    to_page: int | None = None,
) -> str:
    """
    Convert a Word document to Markdown using MarkItDown.
    
    Args:
        input_path: Path to the input Word document
        from_page: Starting page for extraction (optional, requires Word)
        to_page: Ending page for extraction (optional, requires Word)
    
    Returns:
        Path to the generated Markdown file
    """
    if not markitdown_available:
        raise RuntimeError("MarkItDown is not installed. Run: pip install markitdown")
    
    input_path = Path(input_path)
    _validate_input_file(input_path)
    output_path = _get_output_path(input_path)
    
    temp_files: list[str] = []
    
    try:
        work_path: Path = input_path
        original_docx_path: Path = input_path  # Keep track for image extraction
        
        # Handle .doc conversion and page extraction (Windows + Word only)
        needs_word = (
            input_path.suffix.lower() == '.doc' or
            (from_page is not None and to_page is not None)
        )
        
        if needs_word:
            if not word_available:
                if input_path.suffix.lower() == '.doc':
                    raise RuntimeError("Cannot convert .doc files without Microsoft Word. Please save as .docx first.")
                if from_page is not None:
                    logger.warning("Page extraction requires Microsoft Word - converting full document")
            else:
                with WordInstance() as word:
                    if input_path.suffix.lower() == '.doc':
                        logger.info("Converting .doc to .docx...")
                        work_path = Path(word.convert_to_docx(input_path))
                        original_docx_path = work_path
                        temp_files.append(str(work_path))
                    
                    if from_page is not None and to_page is not None:
                        logger.info(f"Extracting pages {from_page}-{to_page}...")
                        work_path = Path(word.extract_pages(work_path, from_page, to_page))
                        original_docx_path = work_path
                        temp_files.append(str(work_path))
        
        # Convert with MarkItDown
        logger.info(f"Converting with MarkItDown...")
        logger.info(f"  Input: {work_path}")
        logger.info(f"  Output: {output_path}")
        
        md = MarkItDown()
        result = md.convert(str(work_path))
        content = result.text_content
        
        # Fix broken image placeholders
        content = fix_markdown_images(content, original_docx_path, output_path)
        
        # Write output
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)
        
        # Clean up empty media folder if no images were extracted
        media_dir = output_path.parent / MEDIA_FOLDER_NAME
        remove_empty_folder(media_dir)
        
        return str(output_path)
    
    finally:
        for tmp in temp_files:
            try:
                os.unlink(tmp)
            except OSError as e:
                logger.warning(f"Could not remove temp file {tmp}: {e}")

# =================== GUI ===================
def create_gui():
    """Create and run the GUI application."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    
    class ConverterGUI:
        def __init__(self, root: tk.Tk):
            self.root = root
            self.root.title("Word to Markdown Converter")
            self.root.config(bg="#252526")
            self.root.option_add("*Font", FONT_UI)
            self.root.minsize(550, 350)
            
            self.file_var = tk.StringVar()
            self.enable_page_extraction_var = tk.BooleanVar(value=False)
            self.current_pages: int | None = None
            self._check_timer: str | None = None
            
            self._create_widgets()
            self._update_status()
        
        def _create_widgets(self):
            main_frame = tk.Frame(self.root, bg="#252526")
            main_frame.pack(expand=True, fill="both", padx=25, pady=20)
            
            self._create_file_section(main_frame)
            self._create_page_section(main_frame)
            self._create_status_section(main_frame)
            self._create_buttons(main_frame)
        
        def _create_file_section(self, parent):
            tk.Label(
                parent, text="Input File:", fg="#cccccc", bg="#252526",
                font=(FONT_UI[0], FONT_UI[1], "bold")
            ).pack(anchor="w", pady=(0, 5))
            
            entry_frame = tk.Frame(parent, bg="#252526")
            entry_frame.pack(fill="x")
            
            self.file_var.trace_add("write", self._on_file_change)
            self.file_entry = tk.Entry(
                entry_frame, textvariable=self.file_var, font=FONT_MONO,
                bg="#333333", fg="white", insertbackground="white"
            )
            self.file_entry.pack(side="left", fill="x", expand=True, pady=5)
            
            tk.Button(
                entry_frame, text="Browse...", command=self._browse_file,
                bg="#0078d4", fg="white", relief="flat", padx=15
            ).pack(side="left", padx=(10, 0), pady=5)
            
            tk.Label(
                parent, text="Or drag and drop a .docx file onto this script.",
                fg="gray", bg="#252526"
            ).pack(anchor="w", pady=2)
        
        def _create_page_section(self, parent):
            page_frame = tk.LabelFrame(
                parent, text="Page Range (optional - requires MS Word)",
                fg="#cccccc", bg="#252526", relief="groove", bd=1
            )
            page_frame.pack(fill="x", pady=(20, 10))
            
            inner_frame = tk.Frame(page_frame, bg="#252526")
            inner_frame.pack(padx=10, pady=10)
            
            self.page_check = tk.Checkbutton(
                inner_frame, text="Enable Page Extraction",
                variable=self.enable_page_extraction_var,
                command=self._toggle_page_extraction,
                fg="#cccccc", bg="#252526", selectcolor="#333333",
                activebackground="#252526", activeforeground="#cccccc",
                state="normal" if word_available else "disabled"
            )
            self.page_check.pack(side='left', padx=5)
            
            tk.Label(inner_frame, text="From:", fg="#cccccc", bg="#252526").pack(side='left', padx=(20, 5))
            self.from_spin = ttk.Spinbox(inner_frame, from_=1, to=9999, width=6, state='disabled')
            self.from_spin.set(1)
            self.from_spin.pack(side='left', padx=5)
            
            tk.Label(inner_frame, text="To:", fg="#cccccc", bg="#252526").pack(side='left', padx=5)
            self.to_spin = ttk.Spinbox(inner_frame, from_=1, to=9999, width=6, state='disabled')
            self.to_spin.set(1)
            self.to_spin.pack(side='left', padx=5)
            
            if not word_available:
                tk.Label(
                    page_frame, text="⚠ Requires pywin32 and Microsoft Word",
                    fg="#ff9900", bg="#252526"
                ).pack(pady=(0, 5))
        
        def _create_status_section(self, parent):
            status_frame = tk.Frame(parent, bg="#252526")
            status_frame.pack(fill="x", pady=15)
            
            self.status_label = tk.Label(
                status_frame, text="", fg="#888888", bg="#252526"
            )
            self.status_label.pack(anchor="w")
        
        def _create_buttons(self, parent):
            button_frame = tk.Frame(parent, bg="#252526")
            button_frame.pack(pady=20)
            
            tk.Button(
                button_frame, text="Convert to Markdown", command=self._convert,
                font=(FONT_UI[0], 12, "bold"), bg="#00b359", fg="white",
                relief="flat", padx=25, pady=10,
                activebackground="#00994C", activeforeground="white"
            ).pack(side='left', padx=10)
            
            tk.Button(
                button_frame, text="Exit", command=self.root.quit,
                bg="#555555", fg="white", relief="flat", padx=25, pady=10,
                activebackground="#444444", activeforeground="white"
            ).pack(side='left', padx=10)
        
        def _update_status(self):
            if markitdown_available:
                self.status_label.config(text="✓ MarkItDown ready", fg="#00cc00")
            else:
                self.status_label.config(text="✗ MarkItDown not installed", fg="#ff6666")
        
        def _browse_file(self):
            filename = filedialog.askopenfilename(
                title="Select Word Document",
                filetypes=[("Word Documents", "*.docx"), ("Old Word Documents", "*.doc"), ("All Files", "*.*")]
            )
            if filename:
                self.file_var.set(filename)
        
        def _on_file_change(self, *args):
            if self._check_timer:
                self.root.after_cancel(self._check_timer)
            if self.enable_page_extraction_var.get():
                self._check_timer = self.root.after(500, self._update_page_controls)
        
        def _toggle_page_extraction(self):
            if self.enable_page_extraction_var.get():
                self._update_page_controls()
            else:
                self.from_spin.config(state='disabled')
                self.to_spin.config(state='disabled')
        
        def _update_page_controls(self):
            filepath_str = self.file_var.get().strip().strip('"')
            if not filepath_str or not word_available:
                return
            
            filepath = Path(filepath_str)
            if not filepath.exists():
                return
            
            self.from_spin.config(state='disabled')
            self.to_spin.config(state='disabled')
            self.status_label.config(text="Detecting page count...", fg="#888888")
            
            def check_pages():
                pages = get_page_count_background(filepath)
                self.root.after(0, lambda: self._update_ui_pages(pages))
            
            threading.Thread(target=check_pages, daemon=True).start()
        
        def _update_ui_pages(self, pages: int | None):
            if pages:
                self.current_pages = pages
                self.from_spin.config(state='normal', to=pages)
                self.to_spin.config(state='normal', to=pages)
                self.to_spin.set(pages)
                self.status_label.config(text=f"✓ Document has {pages} pages", fg="#00cc00")
            else:
                self.status_label.config(text="Could not detect page count", fg="#ff9900")
        
        def _convert(self):
            filepath = self.file_var.get().strip().strip('"')
            if not filepath:
                messagebox.showerror("Error", "Please select a file first!")
                return
            
            if not Path(filepath).exists():
                messagebox.showerror("Error", f"File not found: {filepath}")
                return
            
            if not markitdown_available:
                messagebox.showerror("Error", "MarkItDown is not installed.\nRun: pip install markitdown")
                return
            
            from_p, to_p = None, None
            if self.enable_page_extraction_var.get() and word_available:
                from_p = int(self.from_spin.get())
                to_p = int(self.to_spin.get())
                if from_p == 1 and to_p == self.current_pages:
                    from_p, to_p = None, None
            
            self.status_label.config(text="Converting...", fg="#888888")
            threading.Thread(
                target=self._run_conversion,
                args=(filepath, from_p, to_p),
                daemon=True
            ).start()
        
        def _run_conversion(self, filepath: str, from_p: int | None, to_p: int | None):
            with managed_console():
                try:
                    print("=" * 60)
                    print("Word to Markdown Converter")
                    print("=" * 60)
                    if from_p and to_p:
                        print(f"Pages: {from_p} to {to_p}")
                    print("-" * 60)
                    
                    output = convert_file(filepath, from_p, to_p)
                    
                    print("-" * 60)
                    print(f"\n✅ SUCCESS!")
                    print(f"Output: {output}")
                    
                    # Show image count
                    output_path = Path(output)
                    media_dir = output_path.parent / MEDIA_FOLDER_NAME
                    if media_dir.exists():
                        image_count = len(list(media_dir.iterdir()))
                        print(f"Images: {image_count} files in {media_dir}")
                    
                    self.root.after(0, lambda: self.status_label.config(
                        text=f"✓ Saved: {output}", fg="#00cc00"
                    ))
                    
                except Exception as e:
                    print(f"\n❌ ERROR: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    self.root.after(0, lambda: self.status_label.config(
                        text=f"Error: {str(e)}", fg="#ff6666"
                    ))
                
                input("\nPress Enter to close...")
    
    root = tk.Tk()
    ConverterGUI(root)
    root.mainloop()

# =================== QUICK CONVERT (Drag & Drop) ===================
def quick_convert(filepath: str):
    """Simple conversion for drag-and-drop mode."""
    with managed_console():
        try:
            print("=" * 60)
            print("Word to Markdown Converter")
            print("=" * 60)
            
            if not markitdown_available:
                print("\n❌ ERROR: MarkItDown is not installed.")
                print("Run: pip install markitdown")
                input("\nPress Enter to exit...")
                return
            
            print(f"Input: {filepath}")
            print("-" * 60)
            
            output = convert_file(filepath)
            
            print("-" * 60)
            print(f"\n✅ SUCCESS!")
            print(f"Output: {output}")
            
            output_path = Path(output)
            media_dir = output_path.parent / MEDIA_FOLDER_NAME
            if media_dir.exists():
                image_count = len(list(media_dir.iterdir()))
                print(f"Images: {image_count} files in {media_dir}")
            
        except Exception as e:
            print(f"\n❌ ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
        
        input("\nPress Enter to exit...")

# =================== MAIN ENTRY POINT ===================
def main():
    if AUTO_INSTALL_DEPS:
        ensure_all_dependencies()
        
        # Re-check after potential install
        global markitdown_available
        try:
            from markitdown import MarkItDown
            markitdown_available = True
        except ImportError:
            pass
    
    if len(sys.argv) > 1:
        quick_convert(sys.argv[1])
    else:
        create_gui()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        try:
            with managed_console():
                print(f"Fatal Error: {e}")
                import traceback
                traceback.print_exc()
                input("Press Enter to exit...")
        except Exception:
            pass