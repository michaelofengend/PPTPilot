# --- llm_handler.py ---
import json
import os
import openai
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold # For safety settings
import re

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
                            # Store keys in a way that's easy to retrieve
                            if key.strip().upper() == "OPENAI_API_KEY": # Normalize key name
                                API_KEYS["openai_api_key"] = value.strip()
                            elif key.strip().upper() == "GEMINI_API_KEY": # Normalize key name
                                API_KEYS["gemini_api_key"] = value.strip()
            else:
                print(f"Warning: {CREDENTIALS_FILE} not found. API calls will likely fail.")
        except Exception as e:
            print(f"Error loading {CREDENTIALS_FILE}: {e}")
            API_KEYS = {} # Ensure API_KEYS is a dict on error
    return API_KEYS

def _read_xml_file_content(xml_file_path):
    """Reads the content of a single XML file."""
    try:
        with open(xml_file_path, 'r', encoding='utf-8') as f_xml:
            return f_xml.read()
    except Exception as e:
        print(f"Error reading XML file {xml_file_path}: {e}")
        return f"Error reading file: {os.path.basename(xml_file_path)}"

def _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths, image_inputs_present=False):
    """
    Helper function to construct the detailed prompt for the LLM.
    image_inputs_present: Boolean indicating if image data is part of the context for vision models.
    """
    json_summary_for_prompt = json.dumps(ppt_json_data, indent=2)
    if len(json_summary_for_prompt) > 150000: 
        json_summary_for_prompt = (
            f"JSON summary is too large to include fully in this section. "
            f"Total slides: {len(ppt_json_data.get('slides', []))}. "
            f"First slide shapes count: {len(ppt_json_data.get('slides', [{}])[0].get('shapes', [])) if ppt_json_data.get('slides') else 'N/A'}."
            f" (Full JSON was prepared but summarized for this prompt view)"
        )
    
    aggregated_xml_content = "\n\n--- Aggregated XML Content (Multiple Files) ---\n"
    total_xml_chars = 0
    xml_files_processed_count = 0

    for xml_path in xml_file_paths:
        content = _read_xml_file_content(xml_path)
        if len(content) > 50000 and xml_files_processed_count > 5 :
             aggregated_xml_content += f"\n\n--- XML File: {os.path.basename(xml_path)} (Content truncated due to length) ---\n{content[:1000]}...\n--- End of XML File: {os.path.basename(xml_path)} ---\n"
             total_xml_chars += 1000 
        else:
            aggregated_xml_content += f"\n\n--- XML File: {os.path.basename(xml_path)} ---\n{content}\n--- End of XML File: {os.path.basename(xml_path)} ---\n"
            total_xml_chars += len(content)
        xml_files_processed_count +=1
        
        if total_xml_chars > 500000: 
            print("Warning: Total XML content is very large, further XML files will be skipped in the prompt.")
            aggregated_xml_content += "\n\n--- Further XML content truncated due to overall size limit. ---\n"
            break

    # Base prompt text
    prompt_text_parts = [
        f"User's request: {user_prompt}",
        "\nYou are an AI assistant that helps modify PowerPoint presentations by editing their underlying XML structure.",
        "You will be provided with:",
        "1. The user's natural language request.",
        "2. A JSON summary of the presentation's content and structure.",
        "3. The full raw XML content of all constituent files from the .pptx package (e.g., slide1.xml, presentation.xml, theme1.xml, etc.)."
    ]

    if image_inputs_present:
        prompt_text_parts.append("4. Image content from the presentation is also part of the input.")
        # Note: Actual image data is handled by the API call functions, not directly in this text prompt string.

    prompt_text_parts.extend([
        "\nYour task is to:",
        "1. Understand the user's request.",
        "2. Identify which XML file(s) need to be modified to fulfill the request.",
        "3. Generate the **complete, new XML content** for each file that needs to be changed.",
        "\nIMPORTANT INSTRUCTIONS FOR YOUR RESPONSE:",
        "- If you modify one or more XML files, for EACH modified file, you MUST present its new content in the following format:",
        "MODIFIED_XML_FILE: [original_filename_including_internal_path_if_known_e.g.,_ppt/slides/slide1.xml_or_just_slide1.xml]",
        "```xml",
        "[Your complete new XML content for this file here]",
        "```",
        "- If the original internal path (e.g., `ppt/slides/slide1.xml`) is known from the input, use it. If only the base name (e.g., `slide1.xml`) is clear, use that.",
        "- Ensure the XML you provide is well-formed and complete for that file.",
        "- If no XML modifications are needed, or if you cannot fulfill the request, explain why.",
        "- You can provide a brief explanation of the changes you made before presenting the XML blocks.",
        "\nHere is the data:",
        f"\n1. JSON Summary of the Presentation:\n{json_summary_for_prompt}",
        f"\n2. Raw XML Content:\n(This section contains the raw XML content from the .pptx package)\n{aggregated_xml_content}"
    ])
    
    final_prompt_text = "\n".join(prompt_text_parts)

    print(f"Constructed prompt. Approx. JSON length: {len(json_summary_for_prompt)}, Approx. XML content length: {total_xml_chars}")
    if total_xml_chars > 300000: 
        print("WARNING: The aggregated XML content is very large and may exceed LLM token limits or be very costly.")
    return final_prompt_text

def call_openai_api(user_prompt, ppt_json_data, xml_file_paths, model_id="gpt-3.5-turbo", image_inputs=None):
    """
    Calls the OpenAI API.
    model_id: Specific OpenAI model ID (e.g., "gpt-3.5-turbo", "gpt-4o", "gpt-4-turbo").
    image_inputs: List of image data for vision models.
                  Example: [{"type": "image_url", "image_url": {"url": "data:image/jpeg;base64,..."}}]
                  This function assumes image_inputs are already formatted correctly if provided.
    """
    keys = load_api_keys()
    api_key = keys.get("openai_api_key")
    response_data = {"text_response": "", "model_used": model_id}

    if not api_key:
        response_data["text_response"] = f"Error: OpenAI API key not found in {CREDENTIALS_FILE}"
        return response_data

    try:
        client = openai.OpenAI(api_key=api_key)
        
        # Construct messages for the API call
        # For OpenAI, content can be a string (for text-only models) or a list of parts (for multimodal)
        message_content_parts = []
        text_prompt_content = _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths, bool(image_inputs))
        message_content_parts.append({"type": "text", "text": text_prompt_content})

        if image_inputs and model_id in ["gpt-4o", "gpt-4-turbo", "gpt-4-vision-preview"]: # Check if model is vision-capable
            message_content_parts.extend(image_inputs) # image_inputs should be list of {"type": "image_url", ...}
            print(f"--- Calling OpenAI API ({model_id}) with {len(image_inputs)} image(s) ---")
        else:
            if image_inputs:
                 print(f"Warning: Images provided but model {model_id} may not be vision-capable for OpenAI. Sending text only.")
            print(f"--- Calling OpenAI API ({model_id}) (text only) ---")
        
        # If only text, content is a string. If multimodal, content is a list of parts.
        # The OpenAI SDK handles this correctly if `messages_content_parts` is structured as a list of dicts.
        payload_content = message_content_parts if (image_inputs and model_id in ["gpt-4o", "gpt-4-turbo"]) else text_prompt_content


        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": payload_content}],
            model=model_id,
            # max_tokens=4090 # Example, adjust as needed, especially for vision models
        )
        response_data["text_response"] = chat_completion.choices[0].message.content
        print("--- OpenAI API Call Successful ---")
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
    """
    Calls the Google Gemini API.
    model_id: Specific Gemini model ID (e.g., "gemini-1.5-flash-latest", "gemini-1.5-pro-latest").
    image_inputs: List of image data.
                  Example: [{"path": "path/to/image.jpg", "mime_type": "image/jpeg"}]
                  or [{"data": image_bytes, "mime_type": "image/png"}]
    """
    keys = load_api_keys()
    api_key = keys.get("gemini_api_key")
    response_data = {"text_response": "", "model_used": model_id}

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
        
        prompt_parts = []
        text_prompt_content = _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths, bool(image_inputs))
        prompt_parts.append(text_prompt_content)

        if image_inputs: # Gemini 1.5 models are multimodal
            num_images_processed = 0
            for img_input in image_inputs:
                try:
                    if "path" in img_input and os.path.exists(img_input["path"]):
                        with open(img_input["path"], "rb") as f:
                            img_bytes = f.read()
                        prompt_parts.append({"mime_type": img_input["mime_type"], "data": img_bytes})
                        num_images_processed += 1
                    elif "data" in img_input: # if raw bytes are provided
                         prompt_parts.append({"mime_type": img_input["mime_type"], "data": img_input["data"]})
                         num_images_processed += 1
                    else:
                        print(f"Warning: Invalid image input format for Gemini: {img_input}")
                except Exception as e_img:
                    print(f"Error processing image for Gemini ({img_input.get('path', 'bytes_data')}): {e_img}")
                    prompt_parts.append(f"\n[Error processing image: {os.path.basename(img_input.get('path', 'N/A'))}]")
            print(f"--- Calling Gemini API ({model_id}) with {num_images_processed} image(s) ---")
        else:
             print(f"--- Calling Gemini API ({model_id}) (text only) ---")

        response = model.generate_content(prompt_parts) # Send list of parts
        print("--- Gemini API Call Successful ---")

        if hasattr(response, 'text') and response.text: # Simpler access if available
            response_data["text_response"] = response.text
        elif response.candidates and response.candidates[0].content and response.candidates[0].content.parts:
             response_data["text_response"] = "".join(part.text for part in response.candidates[0].content.parts if hasattr(part, "text"))
        else: # Detailed fallback and error checking
            if response.prompt_feedback and response.prompt_feedback.block_reason:
                response_data["text_response"] = f"Gemini API call blocked: {response.prompt_feedback.block_reason_message or response.prompt_feedback.block_reason} (Reason: {response.prompt_feedback.block_reason})"
            else: 
                all_candidates_stopped = True
                for candidate_idx, candidate in enumerate(response.candidates):
                     if candidate.finish_reason != "STOP":
                         all_candidates_stopped = False
                         reason_msg = f"Candidate {candidate_idx} did not finish correctly. Reason: {candidate.finish_reason}."
                         if hasattr(candidate, 'safety_ratings') and candidate.safety_ratings:
                             reason_msg += f" Safety ratings: {candidate.safety_ratings}"
                         response_data["text_response"] = reason_msg
                         break 
                if all_candidates_stopped:
                    response_data["text_response"] = "Gemini API: No text content found in response parts, but not blocked and all candidates finished with STOP."

    except Exception as e: 
        response_data["text_response"] = f"An error occurred with Gemini API: {e}"
    return response_data

def get_llm_response(user_prompt, ppt_json_data, xml_file_paths, engine_or_model_id="gemini-1.5-flash-latest", image_inputs=None):
    """
    Main function to dispatch to the appropriate LLM API based on the model ID.
    engine_or_model_id: A specific model ID (e.g., "gpt-4o", "gemini-1.5-pro-latest").
    image_inputs: Data for vision models (format depends on the API).
    """
    print(f"--- LLM Handler (get_llm_response) Called for: {engine_or_model_id} ---")
    # ... (logging of prompt/data lengths)

    if engine_or_model_id.startswith("gemini"):
        return call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_id=engine_or_model_id, image_inputs=image_inputs)
    elif engine_or_model_id.startswith("gpt"):
        return call_openai_api(user_prompt, ppt_json_data, xml_file_paths, model_id=engine_or_model_id, image_inputs=image_inputs)
    else:
        print(f"Warning: engine_or_model_id '{engine_or_model_id}' not recognized as a specific OpenAI/Gemini model ID. Defaulting to gemini-1.5-flash-latest.")
        return call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_id="gemini-1.5-flash-latest", image_inputs=image_inputs)

def parse_llm_response_for_xml_changes(llm_text_response):
    """
    Parses the LLM's text response to extract modified XML file content.
    """
    modified_files = {}
    pattern = re.compile(r"MODIFIED_XML_FILE:\s*(?P<filename>[^\n`]+?)\s*```xml\n(?P<xml_content>.+?)\n```", re.DOTALL)
    
    for match in pattern.finditer(llm_text_response):
        filename = match.group("filename").strip()
        # Normalize filename: replace backslashes with forward slashes, remove potential leading/trailing quotes
        filename = filename.replace("\\", "/").strip('\'"')
        xml_content = match.group("xml_content").strip()
        modified_files[filename] = xml_content
        print(f"Successfully parsed modified XML for: {filename}")

    if not modified_files:
        print("No 'MODIFIED_XML_FILE:' blocks found in LLM response.")
    return modified_files
