# --- llm_handler.py ---
import json
import os
import openai
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold # For safety settings
import re
import time # <--- Added for timing
from pathlib import Path # Added for Path operations
import base64 # For image encoding
from PIL import Image

# --- Configuration & API Key Loading ---
CREDENTIALS_FILE = "credentials.env" 
API_KEYS = {}

def load_api_keys():
    """Loads API keys from credentials.env"""
    global API_KEYS
    if not API_KEYS: # Load only once
        API_KEYS = {} # Initialize as dict
        try:
            if os.path.exists(CREDENTIALS_FILE):
                with open(CREDENTIALS_FILE, 'r') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#') and '=' in line:
                            key, value = line.split('=', 1)
                            if key.strip().upper() == "OPENAI_API_KEY":
                                API_KEYS["openai_api_key"] = value.strip()
                            elif key.strip().upper() == "GEMINI_API_KEY":
                                API_KEYS["gemini_api_key"] = value.strip()
            else:
                print(f"Warning: {CREDENTIALS_FILE} not found. API calls will likely fail.")
        except Exception as e:
            print(f"Error loading {CREDENTIALS_FILE}: {e}")
            API_KEYS = {} 
    return API_KEYS

def _read_xml_file_content(xml_file_path):
    """Reads the content of a single XML file."""
    try:
        with open(xml_file_path, 'r', encoding='utf-8') as f_xml:
            return f_xml.read()
    except Exception as e:
        print(f"Error reading XML file {xml_file_path}: {e}")
        return f"Error reading file: {Path(xml_file_path).name}"

def _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths, image_inputs_present=False, num_slides_with_images=0):
    """
    Helper function to construct the detailed prompt for the LLM.
    image_inputs_present: Boolean indicating if image data is part of the context for vision models.
    num_slides_with_images: Integer, number of slides for which images are provided.
    """
    json_summary_for_prompt = json.dumps(ppt_json_data, indent=2)
    if len(json_summary_for_prompt) > 150000: 
        json_summary_for_prompt = (
            f"JSON summary is too large to include fully in this section. "
            f"Total slides: {len(ppt_json_data.get('slides', []))}. "
            f"First slide shapes count: {len(ppt_json_data.get('slides', [{}])[0].get('shapes', [])) if ppt_json_data.get('slides') else 'N/A'}."
            f" (Full JSON was prepared but summarized for this prompt view)"
        )
    
    slide_xml_files = sorted(
        [p for p in xml_file_paths if "ppt/slides/slide" in Path(p).as_posix()],
        key=lambda x: int(re.search(r'slide(\d+)\.xml', Path(x).name).group(1)) if re.search(r'slide(\d+)\.xml', Path(x).name) else float('inf')
    )
    other_xml_files = [p for p in xml_file_paths if p not in slide_xml_files]

    per_slide_prompt_parts = ["\n\n--- Per-Slide Information (XML and corresponding Image if provided) ---"]
    
    slide_xml_chars_total = 0
    slides_xml_processed_count = 0

    for slide_xml_path_str in slide_xml_files:
        slide_xml_path_obj = Path(slide_xml_path_str)
        slide_number_match = re.search(r'slide(\d+)\.xml', slide_xml_path_obj.name)
        if not slide_number_match:
            print(f"Warning: Could not determine slide number from filename: {slide_xml_path_obj.name}")
            continue
        
        slide_num_from_filename = int(slide_number_match.group(1))
        slide_xml_content = _read_xml_file_content(slide_xml_path_str)
        
        current_slide_xml_part = f"\n\n--- Slide {slide_num_from_filename} ({slide_xml_path_obj.as_posix()}) ---"
        if image_inputs_present and slide_num_from_filename <= num_slides_with_images:
            current_slide_xml_part += f"\n(An image for Slide {slide_num_from_filename} is provided as part of the multimodal input.)"
        
        if len(slide_xml_content) > 30000:
            slide_xml_display_content = f"{slide_xml_content[:15000]}...\n...{slide_xml_content[-15000:]} (Truncated)"
            slide_xml_chars_total += 30000
        else:
            slide_xml_display_content = slide_xml_content
            slide_xml_chars_total += len(slide_xml_content)

        current_slide_xml_part += f"\nXML Content:\n```xml\n{slide_xml_display_content}\n```"
        per_slide_prompt_parts.append(current_slide_xml_part)
        slides_xml_processed_count += 1
        
        if slide_xml_chars_total > 300000:
            per_slide_prompt_parts.append("\n\n--- Further slide XML content truncated due to overall size limit for slide XMLs. ---")
            break

    aggregated_other_xml_content = "\n\n--- Other Ancillary XML Content (e.g., theme, presentation properties) ---\n"
    total_other_xml_chars = 0
    other_xml_files_processed_count = 0

    for xml_path_str in other_xml_files:
        xml_path_obj = Path(xml_path_str)
        content = _read_xml_file_content(xml_path_str)
        
        if len(content) > 50000 and other_xml_files_processed_count > 3:
             current_other_xml_part = f"\n\n--- XML File: {xml_path_obj.as_posix()} (Content truncated due to length) ---\n{content[:1000]}...\n--- End ---\n"
             total_other_xml_chars += 1000 
        else:
            current_other_xml_part = f"\n\n--- XML File: {xml_path_obj.as_posix()} ---\n{content}\n--- End ---\n"
            total_other_xml_chars += len(content)
        
        aggregated_other_xml_content += current_other_xml_part
        other_xml_files_processed_count +=1
        
        if total_other_xml_chars > 200000:
            aggregated_other_xml_content += "\n\n--- Further ancillary XML content truncated due to overall size limit for other XMLs. ---\n"
            break

    # --- MODIFIED: Restructured the prompt for better LLM adherence ---

    # Part 1: Persona and context setting
    prompt_context_parts = [
        "You are an expert AI assistant that modifies PowerPoint presentations by editing their underlying XML structure. You may also receive images of each slide to provide visual context.",
        "You will now be provided with the complete context for a presentation, which includes:",
        "1. A user's natural language modification request.",
        "2. A JSON summary of the presentation's content.",
        "3. The raw XML content for each slide and other presentation components (like themes, layouts, etc.)."
    ]
    if image_inputs_present:
        prompt_context_parts.append("4. An image of each slide, which will be provided as multimodal input for visual context.")

    # Part 2: The actual data payload
    prompt_data_parts = [
        "\n\n--- PRESENTATION CONTEXT & DATA ---",
        f"\nUser's Request:\n{user_prompt}",
        f"\n\nJSON Summary:\n{json_summary_for_prompt}",
        "".join(per_slide_prompt_parts),
        aggregated_other_xml_content
    ]

    # Part 3: The final, critical instruction set
    prompt_instruction_parts = [
        "\n\n--- TASK & OUTPUT FORMAT ---",
        "\nBased on all the provided context (the user's request, JSON, and all XML files), your task is to identify which XML file(s) must be changed to fulfill the request and generate the complete, new content for each of those files.",
        "\n**CRITICAL: YOUR RESPONSE MUST FOLLOW THESE RULES EXACTLY:**",
        "- If you determine that one or more XML files need to be modified, you MUST format your response by providing each modified file's content within a specific block.",
        "- For EACH modified file, you MUST start the block with the tag `MODIFIED_XML_FILE: [original_filename_e.g.,_ppt/slides/slide1.xml]` followed by the code block.",
        "- Example of the required format for ONE modified file:",
        "MODIFIED_XML_FILE: ppt/slides/slide1.xml",
        "```xml",
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>",
        "<p:sld ...>",
        "  ",
        "</p:sld>",
        "```",
        "- The XML you provide MUST be complete and well-formed for that specific file.",
        "- Use the exact internal file path (e.g., `ppt/slides/slide1.xml`, `ppt/theme/theme1.xml`) as seen in the context above.",
        "- **DO NOT** include any extra conversation, commentary, or explanations outside of the `MODIFIED_XML_FILE:` blocks. If no changes are needed, simply respond with 'No changes needed.'."
    ]
    
    # Combine all parts
    final_prompt_parts = prompt_context_parts + prompt_data_parts + prompt_instruction_parts
    final_prompt_text = "\n".join(final_prompt_parts)

    print(f"Constructed prompt. Approx. JSON length: {len(json_summary_for_prompt)}, Approx. Slide XMLs length: {slide_xml_chars_total}, Approx. Other XMLs length: {total_other_xml_chars}")
    if (slide_xml_chars_total + total_other_xml_chars) > 400000: 
        print("WARNING: The total XML content is very large and may exceed LLM token limits or be very costly.")
    return final_prompt_text

def call_openai_api(user_prompt, ppt_json_data, xml_file_paths, model_id="gpt-3.5-turbo", image_inputs=None):
    keys = load_api_keys()
    api_key = keys.get("openai_api_key")
    response_data = {"text_response": "", "model_used": model_id, "inference_time_seconds": None}

    if not api_key:
        response_data["text_response"] = f"Error: OpenAI API key not found in {CREDENTIALS_FILE}"
        return response_data

    try:
        client = openai.OpenAI(api_key=api_key)
        
        message_content_parts = []
        num_slides_with_images = len(image_inputs) if image_inputs else 0

        text_prompt_content = _construct_llm_input_prompt(
            user_prompt, ppt_json_data, xml_file_paths, 
            bool(image_inputs), num_slides_with_images=num_slides_with_images
        )
        message_content_parts.append({"type": "text", "text": text_prompt_content})

        if image_inputs and model_id in ["gpt-4o", "gpt-4-turbo", "gpt-4-vision-preview"]:
            print(f"--- Preparing {len(image_inputs)} image(s) for OpenAI API ({model_id}) ---")
            for img_data in image_inputs: 
                try:
                    with open(img_data["path"], "rb") as image_file:
                        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                    data_url = f"data:{img_data['mime_type']};base64,{encoded_string}"
                    message_content_parts.append({
                        "type": "image_url",
                        "image_url": {"url": data_url, "detail": "low"} 
                    })
                except Exception as e_img:
                    print(f"Error processing image {img_data['path']} for OpenAI: {e_img}")
                    message_content_parts.append({"type": "text", "text": f"[Error processing image: {Path(img_data['path']).name}]"})
        elif image_inputs:
            print(f"Warning: Images provided but model {model_id} may not be vision-capable for OpenAI. Sending text only.")
        
        payload_content = message_content_parts if (image_inputs and model_id in ["gpt-4o", "gpt-4-turbo", "gpt-4-vision-preview"]) else text_prompt_content

        print(f"--- Calling OpenAI API ({model_id}) (multimodal: {bool(image_inputs and model_id in ['gpt-4o', 'gpt-4-turbo'])}) ---")
        llm_start_time = time.time()
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": payload_content}],
            model=model_id,
        )
        llm_end_time = time.time()
        response_data["inference_time_seconds"] = round(llm_end_time - llm_start_time, 3)
        
        response_data["text_response"] = chat_completion.choices[0].message.content
        print(f"--- OpenAI API Call Successful (took {response_data['inference_time_seconds']:.3f}s) ---")
    except openai.APIConnectionError as e:
        response_data["text_response"] = f"OpenAI API Connection Error: {e}"
    except openai.RateLimitError as e:
        response_data["text_response"] = f"OpenAI API Rate Limit Error: {e}"
    except openai.AuthenticationError as e:
        response_data["text_response"] = f"OpenAI API Authentication Error: {e} (Check your API key)"
    except openai.BadRequestError as e: 
         response_data["text_response"] = f"OpenAI API BadRequestError: {e}. The prompt or image data might be too long or invalid."
    except openai.APIError as e: 
        response_data["text_response"] = f"OpenAI API Error: {e}"
    except Exception as e: 
        response_data["text_response"] = f"An unexpected error occurred with OpenAI API: {e}"
    return response_data


def call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_id="gemini-1.5-flash-latest", image_inputs=None):
    keys = load_api_keys()
    api_key = keys.get("gemini_api_key")
    response_data = {"text_response": "", "model_used": model_id, "inference_time_seconds": None}

    if not api_key:
        response_data["text_response"] = f"Error: Gemini API key not found in {CREDENTIALS_FILE}"
        return response_data

    try:
        genai.configure(api_key=api_key)
        safety_settings = [
            {"category": HarmCategory.HARM_CATEGORY_HARASSMENT, "threshold": HarmBlockThreshold.BLOCK_NONE},
            {"category": HarmCategory.HARM_CATEGORY_HATE_SPEECH, "threshold": HarmBlockThreshold.BLOCK_NONE},
            {"category": HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, "threshold": HarmBlockThreshold.BLOCK_NONE},
            {"category": HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, "threshold": HarmBlockThreshold.BLOCK_NONE},
        ]
        model = genai.GenerativeModel(model_id, safety_settings=safety_settings)
        
        prompt_parts_for_api = []
        num_slides_with_images = len(image_inputs) if image_inputs else 0

        text_prompt_content = _construct_llm_input_prompt(
            user_prompt, ppt_json_data, xml_file_paths, 
            bool(image_inputs), num_slides_with_images=num_slides_with_images
        )
        prompt_parts_for_api.append(text_prompt_content)

        if image_inputs:
            num_images_processed = 0
            for img_data in image_inputs:
                try:
                    if "path" in img_data and os.path.exists(img_data["path"]):
                        with open(img_data["path"], "rb") as f:
                            img_bytes = f.read()
                        prompt_parts_for_api.append({"mime_type": img_data["mime_type"], "data": img_bytes})
                        num_images_processed += 1
                    elif "data" in img_data:
                         prompt_parts_for_api.append({"mime_type": img_data["mime_type"], "data": img_data["data"]})
                         num_images_processed += 1
                    else:
                        print(f"Warning: Invalid image input format for Gemini: {img_data}")
                except Exception as e_img:
                    print(f"Error processing image for Gemini ({img_data.get('path', 'bytes_data')}): {e_img}")
                    prompt_parts_for_api.append(f"\n[Error processing image: {Path(img_data.get('path', 'N/A')).name}]")
            print(f"--- Calling Gemini API ({model_id}) with {num_images_processed} image(s) ---")
        else:
             print(f"--- Calling Gemini API ({model_id}) (text only) ---")

        llm_start_time = time.time()
        response = model.generate_content(prompt_parts_for_api)
        llm_end_time = time.time()
        response_data["inference_time_seconds"] = round(llm_end_time - llm_start_time, 3)
        
        print(f"--- Gemini API Call Successful (took {response_data['inference_time_seconds']:.3f}s) ---")

        if hasattr(response, 'text') and response.text:
            response_data["text_response"] = response.text
        elif response.candidates and response.candidates[0].content and response.candidates[0].content.parts:
             response_data["text_response"] = "".join(part.text for part in response.candidates[0].content.parts if hasattr(part, "text"))
        else:
            if response.prompt_feedback and response.prompt_feedback.block_reason:
                reason_msg = f"Gemini API call blocked. Reason: {response.prompt_feedback.block_reason_message or response.prompt_feedback.block_reason}"
                response_data["text_response"] = reason_msg
            else: 
                response_data["text_response"] = "Gemini API: No text content found in response, and not explicitly blocked."
    except Exception as e: 
        response_data["text_response"] = f"An error occurred with Gemini API: {e}"
    return response_data

def get_llm_response(user_prompt, ppt_json_data, xml_file_paths, engine_or_model_id="gemini-1.5-flash-latest", image_inputs=None):
    print(f"--- LLM Handler (get_llm_response) Called for: {engine_or_model_id} ---")
    
    is_vision_model_family = "gpt-4o" in engine_or_model_id or \
                             "gpt-4-turbo" in engine_or_model_id or \
                             "vision" in engine_or_model_id or \
                             "gemini-1.5" in engine_or_model_id or \
                             "gemini-2.0-flash-preview-image-generation" in engine_or_model_id or \
                             "gemini-2.5" in engine_or_model_id

    actual_image_inputs_to_send = image_inputs if is_vision_model_family else None
    if image_inputs and not is_vision_model_family:
        print(f"Warning: Images provided, but selected model '{engine_or_model_id}' is not recognized as vision-capable. Images will not be sent.")


    if engine_or_model_id.startswith("gemini"):
        return call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_id=engine_or_model_id, image_inputs=actual_image_inputs_to_send)
    elif engine_or_model_id.startswith("gpt"):
        return call_openai_api(user_prompt, ppt_json_data, xml_file_paths, model_id=engine_or_model_id, image_inputs=actual_image_inputs_to_send)
    else:
        print(f"Warning: engine_or_model_id '{engine_or_model_id}' not recognized. Defaulting to gemini-1.5-flash-latest.")
        return call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_id="gemini-1.5-flash-latest", image_inputs=actual_image_inputs_to_send)


def parse_llm_response_for_xml_changes(llm_text_response):
    modified_files = {}
    pattern = re.compile(
        r"MODIFIED_XML_FILE:\s*(?P<filename>[a-zA-Z0-9./\-_]+?\.xml)\s*```xml\n(?P<xml_content>.+?)\n```", 
        re.DOTALL
    )
    
    for match in pattern.finditer(llm_text_response):
        filename = match.group("filename").strip()
        filename = filename.replace("\\", "/").strip('\'"')
        xml_content = match.group("xml_content").strip()
        modified_files[filename] = xml_content
        print(f"Successfully parsed modified XML for: {filename}")

    if not modified_files:
        print("No 'MODIFIED_XML_FILE:' blocks found in LLM response.")
    return modified_files

def call_llm_judge(instruction: str, before_img_path: str, after_img_path: str,
                   before_xml_dict: dict, after_xml_dict: dict,
                   model_id: str = "gemini-2.5-flash-preview-05-20"):
    """
    Calls Gemini to act as a judge, comparing slide images and their underlying XML.
    """
    keys = load_api_keys()
    api_key = keys.get("gemini_api_key")
    if not api_key:
        return {"error": "Gemini API key not found."}

    genai.configure(api_key=api_key)

    xml_diff_prompt_part = ""
    if before_xml_dict and after_xml_dict:
        xml_diff_prompt_part += "\\n\\nAdditionally, here are the changes to the underlying XML files. Use this to verify code-level correctness.\\n"
        for filename, before_content in before_xml_dict.items():
            after_content = after_xml_dict.get(filename, "")
            if len(before_content) > 4000:
                before_content = before_content[:2000] + "\\n...\\n" + before_content[-2000:]
            if len(after_content) > 4000:
                after_content = after_content[:2000] + "\\n...\\n" + after_content[-2000:]

            xml_diff_prompt_part += f"\\n--- XML DIFF FOR: {filename} ---\\n"
            xml_diff_prompt_part += f"--- ORIGINAL XML ---\\n```xml\\n{before_content}\\n```\\n"
            xml_diff_prompt_part += f"--- EDITED XML ---\\n```xml\\n{after_content}\\n```\\n"
            xml_diff_prompt_part += "---------------------------------\\n"


    judge_prompt = """
You are an expert slide-editing judge.

TASK
- Compare the ORIGINAL slide with the EDITED slide, considering both the visual appearance and the underlying XML changes provided.
- Evaluate how well the EDITED slide handles the INSTRUCTION.
- Score for instruction following, and quality of Text, Image, Layout, and Color.

SCORING
Return a valid JSON object with exactly these keys:
{
  "instruction_following": <score>,
  "text_quality": <score>,
  "image_quality": <score>,
  "layout_quality": <score>,
  "color_quality": <score>
}

GUIDELINES
Score each category from 0 to 5 based on the following rubric:

INSTRUCTION_FOLLOWING:
5 = Perfect: All aspects of the instruction were met completely and accurately.
4 = Mostly follows: The main goal was achieved, but minor details from the instruction were missed.
3 = Partially follows: Some parts of the instruction were addressed, but key elements were ignored or incorrect.
2 = Loosely follows: The edit is related to the instruction but fails to address the core request.
1 = Attempted but incorrect: An attempt was made, but it fundamentally misunderstands the instruction.
0 = Does not follow: The edit ignores the instruction entirely.

TEXT_QUALITY:
5 = Perfect: Text content, formatting, and typography are flawless and fully satisfy the instruction.
4 = Mostly correct: Text elements are clearly improved but have minor issues in content, formatting, or typography.
3 = Partially correct: Text improvements are noticeable but have significant issues in content, formatting, or typography.
2 = Slightly changed but inadequate: Some text edits are present but insufficient or poorly implemented.
1 = Attempted but incorrect: Text changes are visible but do not match the instruction or improve the slide.
0 = Completely fails: No meaningful text improvements or changes are severely detrimental.

IMAGE_QUALITY:
5 = Perfect: Images are optimal in selection, placement, sizing, and enhancement, fully satisfying the instruction.
4 = Mostly correct: Images are well-selected and implemented with only minor issues in placement, sizing, or visual quality.
3 = Partially correct: Image improvements are noticeable but have significant issues in selection, placement, sizing, or quality.
2 = Slightly changed but inadequate: Some image edits are present but insufficient or poorly implemented.
1 = Attempted but incorrect: Image changes are visible but do not match the instruction or improve the slide.
0 = Completely fails: No meaningful image improvements or changes are severely detrimental.

LAYOUT_QUALITY:
5 = Perfect: Slide organization, spacing, alignment, and element relationships are flawless and fully satisfy the instruction.
4 = Mostly correct: Layout is clearly improved but has minor issues in organization, spacing, or alignment.
3 = Partially correct: Layout improvements are noticeable but have significant issues in organization, spacing, or alignment.
2 = Slightly changed but inadequate: Some layout edits are present but insufficient or poorly implemented.
1 = Attempted but incorrect: Layout changes are visible but do not match the instruction or improve the slide.
0 = Completely fails: No meaningful layout improvements or changes are severely detrimental.

COLOR_QUALITY:
5 = Perfect: Color scheme, contrast, balance, and emphasis are flawless and fully satisfy the instruction.
4 = Mostly correct: Color choices are clearly improved but have minor issues in scheme, contrast, or emphasis.
3 = Partially correct: Color improvements are noticeable but have significant issues in scheme, contrast, or emphasis.
2 = Slightly changed but inadequate: Some color edits are present but insufficient or poorly implemented.
1 = Attempted but incorrect: Color changes are visible but do not match the instruction or improve the slide.
0 = Completely fails: No meaningful color improvements or changes are severely detrimental.

Judge only what you can see in the provided images.
Return *only* the JSON object and nothing else.
"""
    try:
        before_image = Image.open(before_img_path)
        after_image = Image.open(after_img_path)
        
        generation_config = {
          "temperature": 0.2,
          "response_mime_type": "application/json",
        }
        safety_settings = [
            {"category": HarmCategory.HARM_CATEGORY_HARASSMENT, "threshold": HarmBlockThreshold.BLOCK_NONE},
            {"category": HarmCategory.HARM_CATEGORY_HATE_SPEECH, "threshold": HarmBlockThreshold.BLOCK_NONE},
            {"category": HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, "threshold": HarmBlockThreshold.BLOCK_NONE},
            {"category": HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, "threshold": HarmBlockThreshold.BLOCK_NONE},
        ]
        
        model = genai.GenerativeModel(
            model_name=model_id,
            generation_config=generation_config,
            safety_settings=safety_settings
        )
        
        prompt_parts = [
            "INSTRUCTION: ", instruction, "\\n\\n",
            "ORIGINAL slide:", before_image, "\\n\\n",
            "EDITED slide:", after_image,
            xml_diff_prompt_part, # <-- ADD THE XML DIFFS TO THE PROMPT
            "\\n\\n",
            judge_prompt,
        ]

        response = model.generate_content(prompt_parts)
        return json.loads(response.text)

    except Exception as e:
        print(f"Error calling Gemini judge: {e}")
        return {"error": str(e)}