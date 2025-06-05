# app.py
import os
import json
import shutil
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
import ppt_processor
import llm_handler 
import re 
from pathlib import Path 

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
EXTRACTED_XML_FOLDER = 'extracted_xml_original'
MODIFIED_PPTX_FOLDER = 'modified_ppts'
GENERATED_PDFS_FOLDER = 'generated_pdfs'

ALLOWED_EXTENSIONS = {'pptx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXTRACTED_XML_FOLDER'] = EXTRACTED_XML_FOLDER
app.config['MODIFIED_PPTX_FOLDER'] = MODIFIED_PPTX_FOLDER
app.config['GENERATED_PDFS_FOLDER'] = GENERATED_PDFS_FOLDER


for folder_key in ['UPLOAD_FOLDER', 'EXTRACTED_XML_FOLDER', 'MODIFIED_PPTX_FOLDER', 'GENERATED_PDFS_FOLDER']:
    os.makedirs(app.config[folder_key], exist_ok=True)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download_original/<filename>')
def download_original_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/download_modified/<filename>')
def download_modified_file(filename):
    return send_from_directory(app.config['MODIFIED_PPTX_FOLDER'], filename, as_attachment=True)

@app.route('/view_pdf/<pdf_filename>')
def view_pdf(pdf_filename):
    return send_from_directory(app.config['GENERATED_PDFS_FOLDER'], pdf_filename)


@app.route('/process_ppt', methods=['POST'])
def process_ppt_route():
    if 'ppt_file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['ppt_file']
    prompt_text = request.form.get('prompt', '')
    # Get the specific model ID from the form
    selected_model_id = request.form.get('llm_engine', 'gemini-1.5-flash-latest') # Default if not provided

    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and allowed_file(file.filename):
        original_filename_secure = secure_filename(file.filename)
        original_filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename_secure)
        file.save(original_filepath)

        original_xml_output_dir = os.path.join(app.config['EXTRACTED_XML_FOLDER'], original_filename_secure + "_xml")
        
        if os.path.exists(original_xml_output_dir):
            shutil.rmtree(original_xml_output_dir)
        os.makedirs(original_xml_output_dir, exist_ok=True)
            
        modified_pptx_download_url = None
        modified_xml_contents_for_display = {}
        original_pdf_url = None
        modified_pdf_url = None
        image_inputs_for_llm = None # Initialize

        try:
            json_data = ppt_processor.pptx_to_json(original_filepath)
            extracted_original_xml_full_paths = ppt_processor.extract_xml_from_pptx(original_filepath, original_xml_output_dir)
            
            original_pdf_path = ppt_processor.convert_pptx_to_pdf(original_filepath, app.config['GENERATED_PDFS_FOLDER'])
            if original_pdf_path:
                original_pdf_url = f"/view_pdf/{os.path.basename(original_pdf_path)}"

            xml_paths_for_llm_prompt = []
            for full_path in extracted_original_xml_full_paths:
                relative_path = os.path.relpath(full_path, original_xml_output_dir)
                xml_paths_for_llm_prompt.append(relative_path.replace(os.sep, '/')) 

            # --- Placeholder for image extraction if using vision models ---
            # if "gpt-4o" in selected_model_id or "gemini-1.5" in selected_model_id: # Basic check
            #     # You would need a function in ppt_processor.py to extract images
            #     # and format them correctly for the respective API.
            #     # image_inputs_for_llm = ppt_processor.extract_images_for_llm(original_filepath, selected_model_id)
            #     print(f"Note: Model {selected_model_id} is vision capable, but image extraction from PPTX is not yet implemented.")
            #     pass


            llm_result = llm_handler.get_llm_response(
                user_prompt=prompt_text,
                ppt_json_data=json_data,
                xml_file_paths=extracted_original_xml_full_paths, 
                engine_or_model_id=selected_model_id, # Pass the specific model ID
                image_inputs=image_inputs_for_llm # Pass None for now
            )
            llm_response_text = llm_result.get("text_response", "Error: No text response from LLM.")
            actual_model_used = llm_result.get("model_used", selected_model_id)

            parsed_modified_xml_map = llm_handler.parse_llm_response_for_xml_changes(llm_response_text)
            xml_updates_for_new_pptx = {}

            if parsed_modified_xml_map:
                print(f"LLM suggested changes for: {list(parsed_modified_xml_map.keys())}")
                for llm_filename, new_xml_content in parsed_modified_xml_map.items():
                    normalized_llm_filename = llm_filename.replace("\\", "/").strip('\'"') # Normalize from LLM
                    found_path = None
                    if normalized_llm_filename in xml_paths_for_llm_prompt:
                        found_path = normalized_llm_filename
                    else: 
                        for p in xml_paths_for_llm_prompt:
                            if os.path.basename(p) == os.path.basename(normalized_llm_filename): 
                                found_path = p
                                break
                    if found_path:
                        xml_updates_for_new_pptx[found_path] = new_xml_content
                        modified_xml_contents_for_display[found_path] = new_xml_content
                    else:
                        print(f"Warning: Could not map LLM-specified file '{llm_filename}' (normalized: '{normalized_llm_filename}') to an original XML path.")


                if xml_updates_for_new_pptx:
                    modified_pptx_filename_secure = f"modified_{original_filename_secure}"
                    modified_pptx_filepath = os.path.join(app.config['MODIFIED_PPTX_FOLDER'], modified_pptx_filename_secure)
                    
                    success_creating_modified = ppt_processor.create_modified_pptx(
                        original_filepath,
                        xml_updates_for_new_pptx,
                        modified_pptx_filepath
                    )
                    if success_creating_modified:
                        modified_pptx_download_url = f"/download_modified/{modified_pptx_filename_secure}"
                        modified_pdf_path = ppt_processor.convert_pptx_to_pdf(modified_pptx_filepath, app.config['GENERATED_PDFS_FOLDER'])
                        if modified_pdf_path:
                            modified_pdf_url = f"/view_pdf/{os.path.basename(modified_pdf_path)}"
                    else:
                        print("Failed to create modified PPTX.")
            
            original_pptx_download_url = f"/download_original/{original_filename_secure}"

            return jsonify({
                "message": "File processed successfully",
                "json_data": json_data,
                "xml_files": [os.path.basename(f) for f in extracted_original_xml_full_paths], 
                "llm_response": llm_response_text,
                "llm_engine_used": actual_model_used,
                "original_pptx_download_url": original_pptx_download_url,
                "modified_pptx_download_url": modified_pptx_download_url,
                "modified_xml_data": modified_xml_contents_for_display,
                "original_pdf_url": original_pdf_url, 
                "modified_pdf_url": modified_pdf_url  
            }), 200

        except Exception as e:
            app.logger.error(f"Error processing file: {e}", exc_info=True)
            return jsonify({"error": f"An error occurred during processing: {str(e)}"}), 500
            
    else:
        return jsonify({"error": "File type not allowed"}), 400

if __name__ == '__main__':
    app.run(debug=True)
