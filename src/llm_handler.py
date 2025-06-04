# --- llm_handler.py ---
import json
import os
import openai
import google.generativeai as genai
import re

# --- Configuration & API Key Loading ---
CREDENTIALS_FILE = "credentials.env" # Changed from .yml to .env
API_KEYS = {}

def load_api_keys():
    """Loads API keys from credentials.env"""
    global API_KEYS
    if not API_KEYS:  # Load only once
        API_KEYS = {} # Initialize as an empty dict
        try:
            if os.path.exists(CREDENTIALS_FILE):
                with open(CREDENTIALS_FILE, 'r') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#') and '=' in line:
                            key, value = line.split('=', 1)
                            # Map .env variable names to the keys expected by the script
                            if key.strip() == "OPENAI_API_KEY":
                                API_KEYS["openai_api_key"] = value.strip()
                            elif key.strip() == "GEMINI_API_KEY":
                                API_KEYS["gemini_api_key"] = value.strip()
                            # Add more mappings here if other keys are introduced
            else:
                print(f"Warning: {CREDENTIALS_FILE} not found. API calls will likely fail.")
        except Exception as e:
            print(f"Error loading {CREDENTIALS_FILE}: {e}")
            API_KEYS = {} # Reset on error
    return API_KEYS

def _read_xml_file_content(xml_file_path):
    """Reads the content of a single XML file."""
    try:
        with open(xml_file_path, 'r', encoding='utf-8') as f_xml:
            return f_xml.read()
    except Exception as e:
        print(f"Error reading XML file {xml_file_path}: {e}")
        return f"Error reading file: {os.path.basename(xml_file_path)}"

def _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths):
    """
    Helper function to construct the detailed prompt for the LLM,
    including the content of all XML files and instructions for returning modified XML.
    """
    # Convert JSON data to a string for the prompt.
    # If the JSON is too large, summarize it to avoid exceeding token limits.
    json_summary_for_prompt = json.dumps(ppt_json_data, indent=2)
    if len(json_summary_for_prompt) > 150000: # Arbitrary limit for summarization
        json_summary_for_prompt = (
            f"JSON summary is too large to include fully in this section. "
            f"Total slides: {len(ppt_json_data.get('slides', []))}. "
            f"First slide shapes count: {len(ppt_json_data.get('slides', [{}])[0].get('shapes', [])) if ppt_json_data.get('slides') else 'N/A'}."
            f" (Full JSON was prepared but summarized for this prompt view)"
        )
    
    # Aggregate XML content from all provided paths.
    # Truncate individual large XML files and the total aggregated content if they exceed limits.
    aggregated_xml_content = "\n\n--- Aggregated XML Content (Multiple Files) ---\n"
    total_xml_chars = 0
    xml_files_processed_count = 0

    for xml_path in xml_file_paths:
        content = _read_xml_file_content(xml_path)
        # Truncate content of individual XML files if too long and many files already processed
        if len(content) > 50000 and xml_files_processed_count > 5 : # Arbitrary limits
             aggregated_xml_content += f"\n\n--- XML File: {os.path.basename(xml_path)} (Content truncated due to length) ---\n{content[:1000]}...\n--- End of XML File: {os.path.basename(xml_path)} ---\n"
             total_xml_chars += 1000 # Approximate length added
        else:
            aggregated_xml_content += f"\n\n--- XML File: {os.path.basename(xml_path)} ---\n{content}\n--- End of XML File: {os.path.basename(xml_path)} ---\n"
            total_xml_chars += len(content)
        xml_files_processed_count +=1
        
        # Stop adding more XML content if the total size becomes too large
        if total_xml_chars > 500000: # Arbitrary overall limit
            print("Warning: Total XML content is very large, further XML files will be skipped in the prompt.")
            aggregated_xml_content += "\n\n--- Further XML content truncated due to overall size limit. ---\n"
            break

    # The main prompt text given to the LLM.
    prompt_text = f"""
User's request: {user_prompt}

You are an AI assistant that helps modify PowerPoint presentations by editing their underlying XML structure.
You will be provided with:
1. The user's natural language request.
2. A JSON summary of the presentation's content and structure.
3. The full raw XML content of all constituent files from the .pptx package (e.g., slide1.xml, presentation.xml, theme1.xml, etc.).

Your task is to:
1. Understand the user's request.
2. Identify which XML file(s) need to be modified to fulfill the request.
3. Generate the **complete, new XML content** for each file that needs to be changed.

IMPORTANT INSTRUCTIONS FOR YOUR RESPONSE:
- If you modify one or more XML files, for EACH modified file, you MUST present its new content in the following format:
MODIFIED_XML_FILE: [original_filename_including_internal_path_if_known_e.g.,_ppt/slides/slide1.xml_or_just_slide1.xml]
```xml
[Your complete new XML content for this file here]
```
- If the original internal path (e.g., `ppt/slides/slide1.xml`) is known from the input, use it. If only the base name (e.g., `slide1.xml`) is clear, use that.
- Ensure the XML you provide is well-formed and complete for that file.
- If no XML modifications are needed, or if you cannot fulfill the request, explain why.
- You can provide a brief explanation of the changes you made before presenting the XML blocks.

Here is the data:

1. JSON Summary of the Presentation:
{json_summary_for_prompt}

2. Raw XML Content:
(This section contains the raw XML content from the .pptx package)
{aggregated_xml_content}

Based on ALL the provided data, please generate your response including any modified XML files in the specified format.
"""
    print(f"Constructed prompt. Approx. JSON length: {len(json_summary_for_prompt)}, Approx. XML content length: {total_xml_chars}")
    if total_xml_chars > 300000: # Warning threshold
        print("WARNING: The aggregated XML content is very large and may exceed LLM token limits or be very costly.")
    return prompt_text

def call_openai_api(user_prompt, ppt_json_data, xml_file_paths, model="gpt-3.5-turbo"):
    """Calls the OpenAI API and returns a dict with text_response and model_used."""
    keys = load_api_keys()
    api_key = keys.get("openai_api_key") # Ensure this key matches what's in load_api_keys
    response_data = {"text_response": "", "model_used": model}

    if not api_key:
        response_data["text_response"] = f"Error: OpenAI API key not found in {CREDENTIALS_FILE}" # Updated error message
        return response_data

    try:
        client = openai.OpenAI(api_key=api_key)
        full_prompt = _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths)
        print(f"--- Calling OpenAI API ({model}) ---")

        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": full_prompt}],
            model=model,
        )
        response_data["text_response"] = chat_completion.choices[0].message.content
        print("--- OpenAI API Call Successful ---")
    # Specific OpenAI error handling
    except openai.APIConnectionError as e:
        response_data["text_response"] = f"OpenAI API Connection Error: {e}"
    except openai.RateLimitError as e:
        response_data["text_response"] = f"OpenAI API Rate Limit Error: {e}"
    except openai.AuthenticationError as e:
        response_data["text_response"] = f"OpenAI API Authentication Error: {e} (Check your API key)"
    except openai.BadRequestError as e: 
         response_data["text_response"] = f"OpenAI API BadRequestError: {e}. The prompt might be too long."
    except openai.APIError as e: # More general OpenAI API error
        response_data["text_response"] = f"OpenAI API Error: {e}"
    except Exception as e: # Catch-all for other unexpected errors
        response_data["text_response"] = f"An unexpected error occurred with OpenAI API: {e}"
    return response_data


def call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_name="gemini-1.5-flash-latest"):
    """Calls the Google Gemini API and returns a dict with text_response and model_used."""
    keys = load_api_keys()
    api_key = keys.get("gemini_api_key") # Ensure this key matches what's in load_api_keys
    response_data = {"text_response": "", "model_used": model_name}

    if not api_key:
        response_data["text_response"] = f"Error: Gemini API key not found in {CREDENTIALS_FILE}" # Updated error message
        return response_data

    try:
        genai.configure(api_key=api_key)
        # Configure safety settings to be less restrictive for this use case.
        # Adjust these as necessary based on content policies and observed behavior.
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]
        model = genai.GenerativeModel(model_name, safety_settings=safety_settings)
        full_prompt = _construct_llm_input_prompt(user_prompt, ppt_json_data, xml_file_paths)
        print(f"--- Calling Gemini API ({model_name}) ---")

        response = model.generate_content(full_prompt)
        print("--- Gemini API Call Successful ---")

        # Extract text from response, handling potential variations in structure
        if hasattr(response, 'text') and response.text:
            response_data["text_response"] = response.text
        elif response.candidates and response.candidates[0].content and response.candidates[0].content.parts:
             response_data["text_response"] = "".join(part.text for part in response.candidates[0].content.parts if hasattr(part, "text"))
        else:
            # Handle cases where the response might be blocked or have an unexpected structure
            if response.prompt_feedback and response.prompt_feedback.block_reason:
                response_data["text_response"] = f"Gemini API call blocked: {response.prompt_feedback.block_reason_message or response.prompt_feedback.block_reason} (Reason: {response.prompt_feedback.block_reason})"
            else: 
                # Check if any candidate indicates a non-STOP finish reason
                for candidate in response.candidates:
                     if candidate.finish_reason != "STOP":
                         response_data["text_response"] = f"Gemini API call did not finish correctly. Reason: {candidate.finish_reason}. Safety ratings: {candidate.safety_ratings if hasattr(candidate, 'safety_ratings') else 'N/A'}"
                         break 
                else: # If loop completed without break, no specific error reason found
                    response_data["text_response"] = "Gemini API response structure changed or no text content found, and not blocked."
    except Exception as e: # Catch-all for other unexpected errors
        response_data["text_response"] = f"An error occurred with Gemini API: {e}"
    return response_data

def get_llm_response(user_prompt, ppt_json_data, xml_file_paths, engine="gemini"):
    """
    Main function to dispatch to the appropriate LLM API.
    Returns a dictionary: {"text_response": "...", "model_used": "..."}
    """
    print(f"--- LLM Handler (get_llm_response) Called for engine: {engine} ---")
    print(f"User Prompt length: {len(user_prompt)}")
    print(f"JSON data length: {len(json.dumps(ppt_json_data))}") # Log length of JSON string
    
    # Calculate and log total size of XML files to be processed
    total_xml_size = 0
    if xml_file_paths: # Ensure xml_file_paths is not None and not empty
        try:
            total_xml_size = sum(os.path.getsize(p) for p in xml_file_paths if os.path.exists(p))
        except TypeError: # Handle if xml_file_paths is not iterable or contains non-paths
            print("Warning: Could not calculate total XML size. xml_file_paths might be invalid.")
            total_xml_size = "N/A" # Indicate size couldn't be determined
            
    print(f"Total size of XML files on disk: {total_xml_size} bytes from {len(xml_file_paths) if xml_file_paths else 0} files.")

    # Dispatch to the chosen LLM API
    if engine == "gemini":
        return call_gemini_api(user_prompt, ppt_json_data, xml_file_paths, model_name="gemini-1.5-flash-latest")
    elif engine == "openai":
        return call_openai_api(user_prompt, ppt_json_data, xml_file_paths, model="gpt-3.5-turbo")
    else:
        # Handle unknown engine choice
        return {"text_response": f"Error: Unknown LLM engine '{engine}'. Choose 'gemini' or 'openai'.", "model_used": "N/A"}

def parse_llm_response_for_xml_changes(llm_text_response):
    """
    Parses the LLM's text response to extract modified XML file content.
    Expects format:
    MODIFIED_XML_FILE: [filename]
    ```xml
    [XML content]
    ```
    Returns a dictionary: {'filename.xml': 'xml_content_string', ...}
    """
    modified_files = {}
    # Regex to find the filename and the XML block.
    # (?P<filename>[^\n`]+?) captures the filename: any characters not newline or backtick, non-greedy.
    # \s* allows for optional whitespace around elements.
    # ```xml\n marks the start of the XML block.
    # (?P<xml_content>.+?) captures the XML content: any characters, non-greedy.
    # \n``` marks the end of the XML block.
    # re.DOTALL allows '.' to match newline characters within the XML content.
    pattern = re.compile(r"MODIFIED_XML_FILE:\s*(?P<filename>[^\n`]+?)\s*```xml\n(?P<xml_content>.+?)\n```", re.DOTALL)
    
    for match in pattern.finditer(llm_text_response):
        filename = match.group("filename").strip()
        xml_content = match.group("xml_content").strip()
        modified_files[filename] = xml_content
        print(f"Successfully parsed modified XML for: {filename}")

    if not modified_files:
        print("No 'MODIFIED_XML_FILE:' blocks found in LLM response.")
    return modified_files