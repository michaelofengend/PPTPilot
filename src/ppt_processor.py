# --- ppt_processor.py ---
from pptx import Presentation
import json
import zipfile
import os
import shutil
from pathlib import Path
import subprocess # For calling LibreOffice

def extract_text_from_shape(shape):
    """Extracts text from a shape, handling different shape types."""
    text = ""
    if shape.has_text_frame:
        text = shape.text_frame.text
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                text += cell.text_frame.text + "\t" 
            text += "\n" 
    return text.strip()

def pptx_to_json(filepath):
    """Converts a .pptx file to a JSON representation."""
    try:
        prs = Presentation(filepath)
        presentation_data = {
            "filename": os.path.basename(filepath),
            "slides": []
        }
        for i, slide in enumerate(prs.slides):
            slide_data = {
                "slide_number": i + 1,
                "shapes": [],
                "notes": ""
            }
            for shape in slide.shapes:
                shape_info = {
                    "name": shape.name,
                    "type": str(shape.shape_type),
                    "text": extract_text_from_shape(shape),
                    "left": shape.left.pt if hasattr(shape, 'left') and shape.left is not None else None,
                    "top": shape.top.pt if hasattr(shape, 'top') and shape.top is not None else None,
                    "width": shape.width.pt if hasattr(shape, 'width') and shape.width is not None else None,
                    "height": shape.height.pt if hasattr(shape, 'height') and shape.height is not None else None,
                }
                slide_data["shapes"].append(shape_info)
            if slide.has_notes_slide:
                notes_slide = slide.notes_slide
                text_frame = notes_slide.notes_text_frame
                slide_data["notes"] = text_frame.text.strip()
            presentation_data["slides"].append(slide_data)
        return presentation_data
    except Exception as e:
        print(f"Error converting {filepath} to JSON: {e}")
        raise

def extract_xml_from_pptx(pptx_filepath, output_folder):
    """
    Extracts all constituent XML files from a .pptx file.
    Returns a list of full paths to the extracted XML files.
    """
    extracted_files_paths = []
    try:
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(pptx_filepath, 'r') as pptx_zip:
            for member_info in pptx_zip.infolist():
                member_name = member_info.filename
                # Ensure we are not trying to extract directory entries if any are listed
                if not member_info.is_dir() and member_name.endswith(('.xml', '.rels')):
                    target_path = os.path.join(output_folder, member_name)
                    # Ensure parent directory for the file exists
                    # (ZipFile.extract() handles this, but good practice if writing manually)
                    os.makedirs(os.path.dirname(target_path), exist_ok=True)
                    
                    # Extract file by file to handle potential path issues in member_name
                    with pptx_zip.open(member_name) as source, open(target_path, "wb") as target:
                        target.write(source.read())
                    extracted_files_paths.append(target_path)
        return extracted_files_paths
    except Exception as e:
        print(f"Error extracting XML from {pptx_filepath}: {e}")
        raise

def create_modified_pptx(original_pptx_path, modified_xml_map, output_pptx_path):
    """
    Creates a new .pptx file by taking an original .pptx, and replacing
    specified internal XML files with new content. This version reads all
    members and writes to a new zip to avoid issues with in-place modification.
    """
    temp_output_pptx_path = output_pptx_path + ".tmp" # Intermediate temporary file

    try:
        with zipfile.ZipFile(original_pptx_path, 'r') as zin:
            with zipfile.ZipFile(temp_output_pptx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    # Normalize the item name for comparison (as keys in modified_xml_map)
                    item_name_normalized = item.filename.replace("\\", "/")
                    
                    if item_name_normalized in modified_xml_map:
                        # If this item is one to be modified, write the new content
                        new_content = modified_xml_map[item_name_normalized]
                        zout.writestr(item, new_content.encode('utf-8'))
                        print(f"Successfully wrote modified '{item_name_normalized}' to temp zip.")
                    else:
                        # Otherwise, copy the item from the original zip
                        buffer = zin.read(item.filename)
                        zout.writestr(item, buffer)
        
        # Replace the target output file with the temporary one
        os.replace(temp_output_pptx_path, output_pptx_path)
        print(f"Modified PPTX successfully created at: {output_pptx_path}")
        return True

    except Exception as e:
        print(f"Error creating modified PPTX at {output_pptx_path}: {e}")
        if os.path.exists(temp_output_pptx_path):
            os.remove(temp_output_pptx_path) # Clean up temp file on error
        return False
    finally:
        # Ensure temp file is removed if it still exists for some reason (e.g. os.replace failed)
        if os.path.exists(temp_output_pptx_path) and not os.path.exists(output_pptx_path) : # only remove if os.replace failed
             if os.path.exists(temp_output_pptx_path): # Re-check existence before removing
                try:
                    os.remove(temp_output_pptx_path)
                except OSError as e_remove:
                    print(f"Warning: Could not remove temporary file {temp_output_pptx_path}: {e_remove}")


def convert_pptx_to_pdf(pptx_filepath, output_folder):
    """
    Converts a .pptx file to .pdf using LibreOffice/OpenOffice.
    Args:
        pptx_filepath (str): Path to the input .pptx file.
        output_folder (str): Folder where the PDF will be saved.
    Returns:
        str: Path to the generated PDF file, or None if conversion failed.
    """
    # Ensure paths are absolute for LibreOffice
    abs_pptx_filepath = os.path.abspath(pptx_filepath)
    abs_output_folder = os.path.abspath(output_folder)

    if not os.path.exists(abs_pptx_filepath):
        print(f"Error: PPTX file not found at {abs_pptx_filepath}")
        return None

    Path(abs_output_folder).mkdir(parents=True, exist_ok=True)
    
    soffice_commands = ['libreoffice', 'soffice']
    soffice_cmd_to_use = None
    
    for cmd_test in soffice_commands:
        try:
            # Check if the command exists and is executable
            # Using shell=True for Windows compatibility if soffice is in PATH but not directly executable
            # For macOS/Linux, direct execution is usually fine.
            # A more robust check involves shutil.which(cmd_test)
            if shutil.which(cmd_test): # Check if command is in PATH and executable
                subprocess.run([cmd_test, '--version'], capture_output=True, check=True, timeout=10)
                soffice_cmd_to_use = cmd_test
                print(f"Found working soffice command: {soffice_cmd_to_use}")
                break 
        except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            print(f"Soffice command '{cmd_test}' not working or timed out: {e}")
            continue
            
    if not soffice_cmd_to_use:
        print("Error: LibreOffice/OpenOffice command ('libreoffice' or 'soffice') not found or not working.")
        print("Please install LibreOffice and ensure it's in your system PATH to enable PDF conversion.")
        return None

    try:
        print(f"Attempting to convert {abs_pptx_filepath} to PDF using {soffice_cmd_to_use} into {abs_output_folder}...")
        # Using a list of arguments is generally safer than a single string with shell=True
        command_args = [
            soffice_cmd_to_use,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', abs_output_folder,
            abs_pptx_filepath
        ]
        
        process = subprocess.run(
            command_args,
            capture_output=True,
            text=True, # Decodes stdout/stderr as text
            timeout=120  # Increased timeout for potentially large files
        )
        
        # LibreOffice usually names the output file the same as input but with .pdf extension
        pdf_filename = Path(abs_pptx_filepath).stem + ".pdf"
        generated_pdf_path = os.path.join(abs_output_folder, pdf_filename)

        if process.returncode == 0 and os.path.exists(generated_pdf_path):
            print(f"Successfully converted to PDF: {generated_pdf_path}")
            return generated_pdf_path
        else:
            print(f"Error during PDF conversion for {abs_pptx_filepath}.")
            print(f"Return code: {process.returncode}")
            print(f"SOFFICE STDOUT: {process.stdout.strip() if process.stdout else 'N/A'}")
            print(f"SOFFICE STDERR: {process.stderr.strip() if process.stderr else 'N/A'}")
            if not os.path.exists(generated_pdf_path):
                print(f"Output PDF file not found at expected location: {generated_pdf_path}")
            return None
            
    except FileNotFoundError: # Should be caught by shutil.which check now
        print(f"Error: '{soffice_cmd_to_use}' command not found. Is LibreOffice/OpenOffice installed and in PATH?")
        return None
    except subprocess.TimeoutExpired:
        print(f"Error: PDF conversion timed out for {abs_pptx_filepath}.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred during PDF conversion: {e}")
        return None

