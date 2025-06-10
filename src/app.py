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
import time
import csv
from datetime import datetime

app = Flask(__name__)

# --- MODIFIED: Configuration ---
# UPLOAD_FOLDER is removed, as we now reference the benchmark ppts directly
SCRIPT_DIR = Path(__file__).parent.resolve()
TSBENCH_PRESENTATIONS_DIR = SCRIPT_DIR / "TSBench" / "benchmark_ppts"
EXTRACTED_XML_FOLDER = SCRIPT_DIR / 'extracted_xml_original'
MODIFIED_PPTX_FOLDER = SCRIPT_DIR / 'modified_ppts'
GENERATED_IMAGES_FOLDER = SCRIPT_DIR / 'generated_images'
PROCESSING_LOG_CSV = SCRIPT_DIR / 'processing_log.csv'


ALLOWED_EXTENSIONS = {'pptx'}

# --- MODIFIED: Use Path objects for consistency ---
app.config['EXTRACTED_XML_FOLDER'] = str(EXTRACTED_XML_FOLDER)
app.config['MODIFIED_PPTX_FOLDER'] = str(MODIFIED_PPTX_FOLDER)
app.config['GENERATED_IMAGES_FOLDER'] = str(GENERATED_IMAGES_FOLDER)
app.config['TSBENCH_PRESENTATIONS_DIR'] = str(TSBENCH_PRESENTATIONS_DIR)

# --- MODIFIED: Create only necessary directories ---
for folder in [EXTRACTED_XML_FOLDER, MODIFIED_PPTX_FOLDER, GENERATED_IMAGES_FOLDER]:
    folder.mkdir(parents=True, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def log_processing_details(log_data):
    """Appends a record to the processing log CSV file."""
    file_exists = os.path.isfile(PROCESSING_LOG_CSV)
    with open(PROCESSING_LOG_CSV, 'a', newline='') as csvfile:
        fieldnames = [
            'Timestamp', 'OriginalFilename', 'LLMEngineUsed', 
            'TotalProcessingTimeSeconds', 'JSONExtractionTimeSeconds', 
            'XMLExtractionTimeSeconds', 'LLMInferenceTimeSeconds', 
            'PPTXModificationTimeSeconds', 'ImageConversionTimeSeconds',
            'TotalSlidesInOriginal', 'NumberOfSlidesEditedByLLM', 
            'ModifiedXMLFilesList'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        if not file_exists:
            writer.writeheader()
        
        writer.writerow(log_data)


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download_modified/<filename>')
def download_modified_file(filename):
    return send_from_directory(app.config['MODIFIED_PPTX_FOLDER'], filename, as_attachment=True)

@app.route('/view_slide_image/<path:image_path>')
def view_slide_image(image_path):
    """Serves an image from the generated_images directory."""
    return send_from_directory(app.config['GENERATED_IMAGES_FOLDER'], image_path, as_attachment=False)


@app.route('/api/process', methods=['POST'])
def process_ppt_route():
    """
    Handles the file upload and processing request from the benchmark runner.
    Provides a more detailed reason when no PPTX file is generated.
    """
    overall_start_time = time.time()
    if 'file' not in request.files:
        return jsonify({"error": "No file part in request. The key should be 'file'."}), 400
    
    file = request.files['file']
    prompt_text = request.form.get('prompt', '')
    selected_model_id = request.form.get('llm_engine', 'gemini-1.5-flash-latest')

    if file.filename == '': return jsonify({"error": "No selected file"}), 400

    original_filename_secure = "N/A"
    try:
        if file and allowed_file(file.filename):
            original_filename_secure = secure_filename(file.filename)
            
            # --- MODIFIED: Construct path to existing benchmark file instead of uploading ---
            original_filepath = os.path.join(app.config['TSBENCH_PRESENTATIONS_DIR'], original_filename_secure)

            if not os.path.exists(original_filepath):
                return jsonify({"error": f"File '{original_filename_secure}' not found in benchmark directory."}), 404
            
            # --- Timing & Processing Steps ---
            time_json_start = time.time()
            json_data = ppt_processor.pptx_to_json(original_filepath)
            time_json_end = time.time()

            time_xml_extract_start = time.time()
            original_xml_output_dir = os.path.join(app.config['EXTRACTED_XML_FOLDER'], original_filename_secure + "_xml")
            if os.path.exists(original_xml_output_dir): shutil.rmtree(original_xml_output_dir)
            extracted_original_xml_full_paths = ppt_processor.extract_xml_from_pptx(original_filepath, original_xml_output_dir)
            time_xml_extract_end = time.time()
            
            xml_paths_for_llm_prompt_relative = [
                Path(p).relative_to(original_xml_output_dir).as_posix() for p in extracted_original_xml_full_paths
            ]
            
            llm_result = llm_handler.get_llm_response(
                user_prompt=prompt_text,
                ppt_json_data=json_data,
                xml_file_paths=extracted_original_xml_full_paths,
                engine_or_model_id=selected_model_id
            )
            actual_model_used = llm_result.get("model_used", selected_model_id)
            parsed_modified_xml_map = llm_handler.parse_llm_response_for_xml_changes(llm_result.get("text_response", ""))
            
            modified_pptx_download_url = None
            edited_slides_comparison_data = []
            number_of_slides_edited = 0
            reason_for_no_modification = None
            time_pptx_modify_start = time_pptx_modify_end = 0
            time_img_conv_start = time_img_conv_end = 0

            if parsed_modified_xml_map:
                xml_updates_for_new_pptx_relative_keys = {}
                edited_slide_numbers = set()

                for llm_filename_key, new_xml_content in parsed_modified_xml_map.items():
                    if llm_filename_key in xml_paths_for_llm_prompt_relative:
                        xml_updates_for_new_pptx_relative_keys[llm_filename_key] = new_xml_content
                        match = re.search(r'ppt/slides/slide(\d+)\.xml', llm_filename_key)
                        if match:
                            edited_slide_numbers.add(int(match.group(1)))
                
                number_of_slides_edited = len(edited_slide_numbers)

                if xml_updates_for_new_pptx_relative_keys:
                    modified_pptx_filename_secure = f"modified_{original_filename_secure}"
                    modified_pptx_filepath = os.path.join(app.config['MODIFIED_PPTX_FOLDER'], modified_pptx_filename_secure)
                    
                    time_pptx_modify_start = time.time()
                    creation_success = ppt_processor.create_modified_pptx(original_filepath, xml_updates_for_new_pptx_relative_keys, modified_pptx_filepath)
                    time_pptx_modify_end = time.time()

                    if creation_success:
                        modified_pptx_download_url = f"/download_modified/{modified_pptx_filename_secure}"

                        time_img_conv_start = time.time()
                        original_img_dir = os.path.join(app.config['GENERATED_IMAGES_FOLDER'], f"{original_filename_secure}_orig")
                        modified_img_dir = os.path.join(app.config['GENERATED_IMAGES_FOLDER'], f"{modified_pptx_filename_secure}_mod")
                        
                        original_image_paths = ppt_processor.export_slides_to_images(original_filepath, original_img_dir)
                        modified_image_paths = ppt_processor.export_slides_to_images(modified_pptx_filepath, modified_img_dir)
                        time_img_conv_end = time.time()
                        
                        abs_generated_images_folder = os.path.abspath(app.config['GENERATED_IMAGES_FOLDER'])

                        for slide_num in sorted(list(edited_slide_numbers)):
                            original_img_path = original_image_paths[slide_num - 1] if len(original_image_paths) >= slide_num else None
                            modified_img_path = modified_image_paths[slide_num - 1] if len(modified_image_paths) >= slide_num else None

                            if original_img_path and modified_img_path:
                                edited_slides_comparison_data.append({
                                    "slide_number": slide_num,
                                    "original_image_url": f"/view_slide_image/{Path(original_img_path).relative_to(abs_generated_images_folder).as_posix()}",
                                    "modified_image_url": f"/view_slide_image/{Path(modified_img_path).relative_to(abs_generated_images_folder).as_posix()}"
                                })
            else:
                # --- MODIFIED: Capture the specific reason for no modification ---
                reason_for_no_modification = "LLM did not return any parsable 'MODIFIED_XML_FILE' blocks."
                llm_text_response = llm_result.get("text_response", "").strip()
                if "no changes needed" in llm_text_response.lower() or len(llm_text_response) < 30:
                    reason_for_no_modification = f"LLM explicitly stated no changes were needed. Full Response: '{llm_text_response}'"
            
            total_processing_time = time.time() - overall_start_time
            timing_stats = {
                "total_processing_time_s": round(total_processing_time, 3),
                "json_extraction_time_s": round(time_json_end - time_json_start, 3),
                "xml_extraction_time_s": round(time_xml_extract_end - time_xml_extract_start, 3),
                "llm_inference_time_s": llm_result.get("inference_time_seconds"),
                "pptx_modification_time_s": round(time_pptx_modify_end - time_pptx_modify_start, 3) if time_pptx_modify_start else "N/A",
                "image_conversion_time_s": round(time_img_conv_end - time_img_conv_start, 3) if time_img_conv_start else "N/A",
                "number_of_slides_edited_by_llm": number_of_slides_edited,
                "total_slides_in_original": len(json_data.get("slides", []))
            }
            
            log_data = {
                'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'OriginalFilename': original_filename_secure,
                'LLMEngineUsed': actual_model_used,
                'TotalProcessingTimeSeconds': timing_stats["total_processing_time_s"],
                'JSONExtractionTimeSeconds': timing_stats["json_extraction_time_s"],
                'XMLExtractionTimeSeconds': timing_stats["xml_extraction_time_s"],
                'LLMInferenceTimeSeconds': timing_stats["llm_inference_time_s"],
                'PPTXModificationTimeSeconds': timing_stats["pptx_modification_time_s"],
                'ImageConversionTimeSeconds': timing_stats["image_conversion_time_s"],
                'TotalSlidesInOriginal': timing_stats["total_slides_in_original"],
                'NumberOfSlidesEditedByLLM': number_of_slides_edited,
                'ModifiedXMLFilesList': ", ".join(parsed_modified_xml_map.keys()) if parsed_modified_xml_map else "None"
            }
            log_processing_details(log_data)

            response_payload = {
                "message": "File processed successfully.",
                "llm_engine_used": actual_model_used,
                "llm_response": llm_result.get("text_response"),
                "modified_pptx_download_url": modified_pptx_download_url,
                "reason_for_no_modification": reason_for_no_modification,
                "edited_slides_comparison_data": edited_slides_comparison_data,
                "timing_stats": timing_stats,
                "json_data": json_data,
                "xml_files": [Path(f).name for f in extracted_original_xml_full_paths],
                "modified_xml_data": parsed_modified_xml_map
            }
            return jsonify(response_payload), 200
        else:
            return jsonify({"error": "File type not allowed"}), 400

    except Exception as e:
        app.logger.error(f"Error processing file '{original_filename_secure}': {e}", exc_info=True)
        return jsonify({"error": f"An error occurred during processing: {str(e)}"}), 500

if __name__ == '__main__':
    # Note: The benchmark runner expects the host to be 127.0.0.1 and port 5001
    app.run(host='127.0.0.1', port=5001, debug=True)