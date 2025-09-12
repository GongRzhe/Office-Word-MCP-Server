"""
Extended document tools for Word Document Server.

These tools provide enhanced document content extraction and search capabilities.
"""
import os
import json
import subprocess
import platform
import shutil
from typing import Dict, List, Optional, Any, Union, Tuple
from docx import Document

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.utils.extended_document_utils import get_paragraph_text, find_text


def get_paragraph_text_from_document(filename: str, paragraph_index: int) -> str:
    """Get text from a specific paragraph in a Word document.
    
    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to retrieve (0-based)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    

    if paragraph_index < 0:
        return "Invalid parameter: paragraph_index must be a non-negative integer"
    
    try:
        result = get_paragraph_text(filename, paragraph_index)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to get paragraph text: {str(e)}"


def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False) -> str:
    """Find occurrences of specific text in a Word document.
    
    Args:
        filename: Path to the Word document
        text_to_find: Text to search for in the document
        match_case: Whether to match case (True) or ignore case (False)
        whole_word: Whether to match whole words only (True) or substrings (False)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    if not text_to_find:
        return "Search text cannot be empty"
    
    try:
        
        result = find_text(filename, text_to_find, match_case, whole_word)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to search for text: {str(e)}"


def convert_to_pdf(filename: str, output_filename: Optional[str] = None) -> str:
    """Convert a Word document to PDF format.
    
    Args:
        filename: Path to the Word document
        output_filename: Optional path for the output PDF. If not provided, 
                         will use the same name with .pdf extension
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Generate output filename if not provided
    if not output_filename:
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}.pdf"
    elif not output_filename.lower().endswith('.pdf'):
        output_filename = f"{output_filename}.pdf"
    
    # Convert to absolute path if not already
    if not os.path.isabs(output_filename):
        output_filename = os.path.abspath(output_filename)
    
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_filename)
    if not output_dir:
        output_dir = os.path.abspath('.')
    
    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Check if output file can be written
    is_writeable, error_message = check_file_writeable(output_filename)
    if not is_writeable:
        return f"Cannot create PDF: {error_message} (Path: {output_filename}, Dir: {output_dir})"
    
    try:
        # Determine platform for appropriate conversion method
        system = platform.system()
        
        if system == "Windows":
            # On Windows, try docx2pdf which uses Microsoft Word
            try:
                from docx2pdf import convert
                convert(filename, output_filename)
                return f"Document successfully converted to PDF: {output_filename}"
            except (ImportError, Exception) as e:
                return f"Failed to convert document to PDF: {str(e)}\nNote: docx2pdf requires Microsoft Word to be installed."
                
        elif system in ["Linux", "Darwin"]:  # Linux or macOS
            # Try using LibreOffice if available (common on Linux/macOS)
            try:
                # Choose the appropriate command based on OS
                if system == "Darwin":  # macOS
                    lo_commands = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
                else:  # Linux
                    lo_commands = ["libreoffice", "soffice"]
                
                # Try each possible command
                conversion_successful = False
                errors = []
                
                for cmd_name in lo_commands:
                    try:
                        # Construct LibreOffice conversion command
                        output_dir = os.path.dirname(output_filename)
                        # If output_dir is empty, use current directory
                        if not output_dir:
                            output_dir = '.'
                        # Ensure the directory exists
                        os.makedirs(output_dir, exist_ok=True)
                        
                        cmd = [
                            cmd_name, 
                            '--headless', 
                            '--convert-to', 
                            'pdf', 
                            '--outdir', 
                            output_dir, 
                            filename
                        ]
                        
                        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                        
                        if result.returncode == 0:
                            # LibreOffice creates the PDF with the same basename
                            base_name = os.path.basename(filename)
                            pdf_base_name = os.path.splitext(base_name)[0] + ".pdf"
                            created_pdf = os.path.join(os.path.dirname(output_filename) or '.', pdf_base_name)
                            
                            # If the created PDF is not at the desired location, move it
                            if created_pdf != output_filename and os.path.exists(created_pdf):
                                shutil.move(created_pdf, output_filename)
                            
                            conversion_successful = True
                            break  # Exit the loop if successful
                        else:
                            errors.append(f"{cmd_name} error: {result.stderr}")
                    except (subprocess.SubprocessError, FileNotFoundError) as e:
                        errors.append(f"{cmd_name} error: {str(e)}")
                
                if conversion_successful:
                    return f"Document successfully converted to PDF: {output_filename}"
                else:
                    # If all LibreOffice attempts failed, try docx2pdf as fallback
                    try:
                        from docx2pdf import convert
                        convert(filename, output_filename)
                        return f"Document successfully converted to PDF: {output_filename}"
                    except (ImportError, Exception) as e:
                        error_msg = "Failed to convert document to PDF using LibreOffice or docx2pdf.\n"
                        error_msg += "LibreOffice errors: " + "; ".join(errors) + "\n"
                        error_msg += f"docx2pdf error: {str(e)}\n"
                        error_msg += "To convert documents to PDF, please install either:\n"
                        error_msg += "1. LibreOffice (recommended for Linux/macOS)\n"
                        error_msg += "2. Microsoft Word (required for docx2pdf on Windows/macOS)"
                        return error_msg
                        
            except Exception as e:
                return f"Failed to convert document to PDF: {str(e)}"
        else:
            return f"PDF conversion not supported on {system} platform"
            
    except Exception as e:
        return f"Failed to convert document to PDF: {str(e)}"
