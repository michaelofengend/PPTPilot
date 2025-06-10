# benchmark_runner.py
import os
import requests
import json
import time
import pandas as pd
from pathlib import Path
from tqdm import tqdm
import shutil
import ppt_processor
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import re
import csv
import logging

# --- Configuration ---
SCRIPT_DIR = Path(__file__).parent.resolve()
PPT_PROCESSOR_URL = "http://127.0.0.1:5001/api/process"
TSBENCH_DIR = SCRIPT_DIR / "tsbench"
TSBENCH_FILE = TSBENCH_DIR / "expanded_instruction_379.json"
TSBENCH_PRESENTATIONS_DIR = TSBENCH_DIR / "benchmark_ppts"
MAX_PROMPTS = 55
# --- NEW: Switchable LLM Engine ---
LLM_ENGINE = "gemini-2.5-flash-preview-05-20"
#LLM_ENGINE = "gpt-4.5-preview-2025-02-27"
# LLM_ENGINE = "o3-2025-04-16"
#LLM_ENGINE = "o1-2025-06-04"
#LLM_ENGINE = "o4-mini"
MAX_CONCURRENT_REQUESTS = 4
REQUEST_TIMEOUT_SECONDS = 300

# --- NEW: Centralized Run Directory ---
RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_OUTPUT_DIR = SCRIPT_DIR / "benchmark_runs" / f"run_{RUN_TIMESTAMP}"
RESULTS_CSV = RUN_OUTPUT_DIR / "benchmark_results.csv"

def process_single_prompt(prompt_id, prompt_text):
    """
    Processes a single prompt: sends request, saves artifacts, and provides
    more detailed error reporting based on the server's response.
    """
    # --- ADDED: Regex to remove the placeholder from the instruction ---
    prompt_text = re.sub(r'\s*\{slide_num\}\s*', '', prompt_text).strip()

    base_id = prompt_id.split('-')[0]
    before_ppt_filename = f"slide_{base_id}.pptx"
    before_ppt_path = TSBENCH_PRESENTATIONS_DIR / before_ppt_filename

    # This is the dedicated directory for all outputs of this single prompt run
    prompt_run_dir = RUN_OUTPUT_DIR / prompt_id
    prompt_run_dir.mkdir(parents=True, exist_ok=True)

    result_entry = {
        "id": prompt_id,
        "instruction": prompt_text,
        "success": False,
        "error_message": "",
        "processing_time_s": None,
        "before_ppt_path": "",
        "output_pptx_path": "",
        "before_images_path": "",
        "after_images_path": "",
        "modified_xml_files": []
    }

    if not before_ppt_path.exists():
        result_entry["error_message"] = f"Skipping: Cannot find 'before' PPTX at {before_ppt_path}"
        return result_entry

    # Copy the 'before' presentation into our run directory for pristine keeping
    shutil.copy(before_ppt_path, prompt_run_dir / "before.pptx")
    result_entry["before_ppt_path"] = str((prompt_run_dir / "before.pptx").relative_to(RUN_OUTPUT_DIR))


    try:
        start_time = time.time()
        with open(before_ppt_path, 'rb') as ppt_file:
            files = {'file': (before_ppt_path.name, ppt_file, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')}
            payload = {'prompt': prompt_text, 'llm_engine': LLM_ENGINE}
            response = requests.post(PPT_PROCESSOR_URL, files=files, data=payload, timeout=REQUEST_TIMEOUT_SECONDS)
        
        result_entry["processing_time_s"] = round(time.time() - start_time, 3)

        if response.ok:
            response_data = response.json()
            result_entry["modified_xml_files"] = list(response_data.get("modified_xml_data", {}).keys())
            
            modified_url = response_data.get("modified_pptx_download_url")
            if modified_url:
                modified_response = requests.get(f"http://127.0.0.1:5001{modified_url}", timeout=REQUEST_TIMEOUT_SECONDS)
                if modified_response.ok:
                    output_path = prompt_run_dir / "after.pptx"
                    with open(output_path, 'wb') as f_out:
                        f_out.write(modified_response.content)
                    
                    result_entry["success"] = True
                    result_entry["output_pptx_path"] = str(output_path.relative_to(RUN_OUTPUT_DIR))

                    # --- Image Generation in the new consolidated structure ---
                    before_img_dir = prompt_run_dir / "before_images"
                    after_img_dir = prompt_run_dir / "after_images"
                    
                    # Generate images for both presentations
                    ppt_processor.export_slides_to_images(str(prompt_run_dir / "before.pptx"), str(before_img_dir))
                    ppt_processor.export_slides_to_images(str(output_path), str(after_img_dir))

                    result_entry["before_images_path"] = str(before_img_dir.relative_to(RUN_OUTPUT_DIR))
                    result_entry["after_images_path"] = str(after_img_dir.relative_to(RUN_OUTPUT_DIR))
                    
                else:
                    # --- MODIFIED: Use the new, more detailed reason from the server ---
                    reason = response_data.get("reason_for_no_modification", "The server returned a response, but downloading the modified PPTX failed.")
                    result_entry["error_message"] = f"PPTX download failed. Reason: {reason}"
            else:
                # --- MODIFIED: Use the new, more detailed reason from the server ---
                reason = response_data.get("reason_for_no_modification", "No modified PPTX URL returned and no reason provided by the server.")
                result_entry["error_message"] = f"Server did not generate a PPTX. Reason: {reason}"
        else:
             try:
                result_entry["error_message"] = response.json().get("error", "Unknown server error.")
             except json.JSONDecodeError:
                result_entry["error_message"] = f"Unknown server error. Status: {response.status_code}, Response: {response.text}"

    except requests.exceptions.RequestException as e:
        result_entry["error_message"] = f"Request failed: {str(e)}"
    except Exception as e:
        result_entry["error_message"] = f"Unexpected error in benchmark runner: {str(e)}"
    
    return result_entry

def run_benchmark():
    """
    Reads the TSBench dataset, sends each entry concurrently to the PPTPilot server,
    and records all results and artifacts in a new timestamped run directory.
    """
    if not TSBENCH_FILE.exists():
        print(f"Error: Benchmark file not found at {TSBENCH_FILE}")
        return

    RUN_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    print(f"Benchmark outputs will be saved to: {RUN_OUTPUT_DIR}")

    with open(TSBENCH_FILE, 'r') as f:
        benchmark_data = json.load(f)

    if not isinstance(benchmark_data, dict):
        print(f"Error: Benchmark file '{TSBENCH_FILE.name}' is not in the expected format.")
        return

    # --- MODIFIED: Create CSV and write header at the start ---
    fieldnames = [
        "id", "instruction", "success", "error_message", "processing_time_s",
        "before_ppt_path", "output_pptx_path", "before_images_path", "after_images_path",
        "modified_xml_files"
    ]
    with open(RESULTS_CSV, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()


    results = []
    
    benchmark_items = list(benchmark_data.items())[:MAX_PROMPTS]

    print(f"Starting benchmark for {len(benchmark_items)} prompts with up to {MAX_CONCURRENT_REQUESTS} parallel requests...")
    print(f"Using LLM Engine: {LLM_ENGINE}")


    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_REQUESTS) as executor:
        # Create a dictionary mapping futures to their prompt_id
        future_to_prompt_id = {executor.submit(process_single_prompt, prompt_id, prompt_text): prompt_id for prompt_id, prompt_text in benchmark_items}
        
        # Process futures as they complete, with a progress bar
        for future in tqdm(as_completed(future_to_prompt_id), total=len(benchmark_items), desc="Processing Prompts"):
            prompt_id = future_to_prompt_id[future]
            try:
                result = future.result()
            except Exception as exc:
                print(f'\nPrompt {prompt_id} generated an exception during execution: {exc}')
                result = {"id": prompt_id, "instruction": benchmark_data[prompt_id], "success": False, "error_message": str(exc)}
            
            # Ensure all fields are present for the CSV writer
            for key in fieldnames:
                if key not in result:
                    result[key] = ""
            
            results.append(result)

            # --- ADDED: Append the result to the CSV immediately ---
            with open(RESULTS_CSV, 'a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writerow(result)


    if not results:
        print("\nBenchmark run finished, but no prompts could be processed.")
        return
        
    # Sort results by the original benchmark order for consistency
    results.sort(key=lambda x: list(benchmark_data.keys()).index(x['id']))


    df = pd.DataFrame(results)
    print(f"\nBenchmark run complete. Results saved incrementally to {RESULTS_CSV}")
    
    # --- Summary calculation ---
    success_count = df['success'].sum()
    total_count = len(df)
    success_rate = (success_count / total_count) * 100 if total_count > 0 else 0
    # Ensure avg_time_col is not empty and contains valid numbers before calculating mean
    avg_time_col = pd.to_numeric(df[df['success'] == True]['processing_time_s'], errors='coerce')
    avg_time = avg_time_col.mean()
    
    print("\n--- Benchmark Summary ---")
    print(f"Total Prompts Attempted: {total_count}")
    print(f"Successful Runs: {success_count}")
    print(f"Success Rate: {success_rate:.2f}%")
    if pd.notna(avg_time):
        print(f"Average Processing Time (for successful runs): {avg_time:.2f}s")
    print("-------------------------")

if __name__ == "__main__":
    run_benchmark()