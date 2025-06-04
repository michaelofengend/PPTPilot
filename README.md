# PPTPilot: AI-Powered PowerPoint Editor

PPTPilot is a Flask-based web application that allows users to upload PowerPoint presentations (`.pptx`), provide natural language prompts, and leverage Large Language Models (LLMs) like OpenAI's GPT or Google's Gemini to modify the presentation's content and structure by directly editing its underlying XML. The application can also convert original and modified presentations to PDF format for easy viewing.

## Features

* **Upload .pptx files:** Securely uploads and stores PowerPoint presentations.
* **LLM Integration:**
    * Supports OpenAI (GPT models like `gpt-3.5-turbo`) and Google Gemini (e.g., `gemini-1.5-flash-latest`).
    * Constructs detailed prompts for LLMs, including a JSON summary of the presentation and the raw XML content of its constituent files.
    * Parses LLM responses to identify and apply changes to specific XML files within the .pptx package.
* **PowerPoint Processing:**
    * Extracts text and structural data from `.pptx` files into a JSON format.
    * Extracts all internal XML files from a `.pptx` package.
    * Creates new `.pptx` files by replacing specified internal XML files with LLM-modified content.
* **PDF Conversion:**
    * Converts original and modified `.pptx` files to `.pdf` format using LibreOffice/OpenOffice.
* **Web Interface:**
    * User-friendly interface built with Flask and Tailwind CSS.
    * Allows users to input a prompt, upload a file, and choose an LLM engine.
    * Displays processing results, including:
        * JSON representation of the original presentation.
        * List of extracted XML filenames.
        * Full LLM text response.
        * LLM engine used.
        * Side-by-side PDF previews of the original and modified presentations.
        * Download links for original and modified `.pptx` files.
        * Details of modified XML content.
    * Provides visual feedback during processing (loader) and displays success/error messages.

## Project Structure

```
PPTPilot/
├── src/
│   ├── app.py                      # Main Flask application, handles routing and UI
│   ├── llm_handler.py              # Handles LLM API calls (OpenAI, Gemini) and response parsing
│   ├── ppt_processor.py            # Core logic for .pptx parsing, XML extraction/modification, PDF conversion
│   ├── credentials.env             # Stores API keys (gitignored by default - CREATE YOUR OWN)
│   ├── templates/
│   │   └── index.html              # HTML template for the web interface
│   ├── uploads/                    # Default folder for uploaded .pptx files
│   ├── extracted_xml_original/     # Stores XML extracted from original .pptx files
│   ├── modified_ppts/              # Stores .pptx files modified by the LLM
│   └── generated_pdfs/             # Stores PDF versions of original and modified .pptx files
└── requirements.txt                # Python dependencies
└── README.md                       # This file
```

## Setup and Installation

1.  **Clone the repository (if applicable):**
    ```bash
    git clone <your-repository-url>
    cd PPTPilot
    ```

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```

3.  **Install Python dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Install LibreOffice/OpenOffice:**
    This application uses LibreOffice (or OpenOffice) to convert PowerPoint files to PDF. Ensure that `libreoffice` or `soffice` command is installed and accessible in your system's PATH.
    * **Linux:** `sudo apt-get install libreoffice` (or similar for your distribution)
    * **macOS:** Download from [LibreOffice website](https://www.libreoffice.org/download/download-libreoffice/) and ensure the command-line tools are in your PATH.
    * **Windows:** Download from [LibreOffice website](https://www.libreoffice.org/download/download-libreoffice/) and ensure its program directory (containing `soffice.exe`) is added to your system's PATH environment variable.

5.  **Create API Credentials File:**
    Create a file named `credentials.env` inside the `src/` directory with the following format:
    ```env
    # src/credentials.env
    OPENAI_API_KEY=sk-YOUR_OPENAI_API_KEY
    GEMINI_API_KEY=YOUR_GEMINI_API_KEY
    ```
    Replace placeholders with your actual API keys. **Important:** Add `src/credentials.env` to your `.gitignore` file if it's not already there to prevent committing your API keys.

## How to Run

1.  Navigate to the `src/` directory:
    ```bash
    cd src
    ```

2.  Run the Flask application:
    ```bash
    flask run
    # Or, for development mode with debugging:
    # python app.py
    ```

3.  Open your web browser and go to `http://127.0.0.1:5000/`.

## Usage

1.  **Enter a prompt:** In the "Your Prompt" text area, describe the changes you want to make to the PowerPoint presentation.
2.  **Upload a PowerPoint file:** Click "Choose File" and select a `.pptx` file.
3.  **Choose LLM Engine:** Select either "Gemini" (default) or "OpenAI (GPT)" from the dropdown.
4.  **Process:** Click the "Process Presentation" button.
5.  **View Results:**
    * A loader will appear while the file is processed.
    * Once complete, the "Processing Results" section will display:
        * PDF previews of the original and (if modified) the new presentation.
        * Download links for the original and modified `.pptx` files.
        * Details about the LLM engine used.
        * The JSON representation of the original PPTX.
        * A list of XML files extracted from the original PPTX.
        * The full response from the LLM.
        * If the LLM made changes, the modified XML content will be shown.

## Key Python Files Overview

* **`app.py`**:
    * Defines Flask routes for the main page (`/`), file processing (`/process_ppt`), downloads (`/download_original/*`, `/download_modified/*`), and PDF viewing (`/view_pdf/*`).
    * Manages file uploads, ensuring they are `.pptx` files.
    * Orchestrates the processing pipeline:
        1.  Saves the uploaded file.
        2.  Calls `ppt_processor.py` functions to convert PPTX to JSON and extract original XML.
        3.  Calls `ppt_processor.py` to convert the original PPTX to PDF.
        4.  Calls `llm_handler.py` to get a response from the chosen LLM.
        5.  Parses the LLM response for XML changes using `llm_handler.py`.
        6.  If changes are suggested, calls `ppt_processor.py` to create a modified `.pptx` file.
        7.  Calls `ppt_processor.py` to convert the modified PPTX to PDF.
        8.  Returns results as JSON to the frontend.
    * Handles error reporting to the user.

* **`llm_handler.py`**:
    * Manages API keys by loading them from `src/credentials.env`.
    * Constructs the detailed prompt sent to the LLM, including the user's request, a JSON summary of the PPT, and the aggregated content of the extracted XML files. It includes logic to truncate oversized JSON or XML content to fit within LLM context limits.
    * Contains functions to call:
        * OpenAI API (`call_openai_api`)
        * Google Gemini API (`call_gemini_api`)
    * Handles API errors gracefully (connection, rate limit, authentication, etc.).
    * Provides a main dispatcher function `get_llm_response` to select the appropriate LLM based on user choice.
    * Includes `parse_llm_response_for_xml_changes` to extract XML content from the LLM's formatted response using regular expressions.

* **`ppt_processor.py`**:
    * `extract_text_from_shape(shape)`: Helper to get text from different PowerPoint shape types (text frames, tables).
    * `pptx_to_json(filepath)`: Converts a `.pptx` file into a structured JSON object containing slide numbers, shapes (with text, type, position, dimensions), and slide notes.
    * `extract_xml_from_pptx(pptx_filepath, output_folder)`: Extracts all constituent XML and `.rels` files from a `.pptx` file (which is a ZIP archive) into a specified output folder.
    * `create_modified_pptx(original_pptx_path, modified_xml_map, output_pptx_path)`: Creates a new `.pptx` file. It takes the original `.pptx`, reads its members, and if a member's path is in `modified_xml_map`, it writes the new content from the map; otherwise, it copies the original member to the new ZIP archive. This avoids in-place modification issues.
    * `convert_pptx_to_pdf(pptx_filepath, output_folder)`: Converts a `.pptx` file to `.pdf` format using an external LibreOffice/OpenOffice command (`soffice` or `libreoffice`). It checks for the presence of these commands and handles potential errors during conversion.

## HTML Frontend (`templates/index.html`)

* A single-page application styled with Tailwind CSS.
* Provides a form for:
    * Text input for the user's prompt.
    * File input for the `.pptx` presentation.
    * Dropdown to select the LLM engine (Gemini or OpenAI).
* JavaScript handles:
    * Form submission via `fetch` API (asynchronous).
    * Displaying a loading indicator during processing.
    * Clearing previous results and messages on new submissions.
    * Rendering results received from the Flask backend:
        * Displaying success or error messages.
        * Populating `pre` tags with JSON data, XML filenames, and LLM responses.
        * Creating download links for original and modified `.pptx` files.
        * Embedding original and modified PDFs in `<iframe>` elements for side-by-side preview.
        * Dynamically showing/hiding sections for PDF previews, download links, and modified XML details based on availability in the server response.
        * Displaying any modified XML content provided by the LLM, organized by filename.

## Potential Improvements / Future Work

* More robust error handling for LLM content generation (e.g., malformed XML).
* Support for more advanced editing features (e.g., adding/deleting slides, image manipulation based on LLM instructions).
* Progress indicators for long-running LLM calls or PDF conversions.
* User authentication and session management if deployed for multiple users.
* Option to select specific slides or elements for the LLM to focus on.
* Directly display XML diffs instead of just the new XML.
* Batch processing of multiple files.
* Caching of LLM responses for identical prompts/files.