# evaluate_results.py
import pandas as pd
from pathlib import Path
import llm_handler # Correctly importing the local module
import ppt_processor
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
import base64
import shutil
import json

# --- Configuration ---
SCRIPT_DIR = Path(__file__).parent.resolve()
BENCHMARK_RUNS_DIR = SCRIPT_DIR / "benchmark_runs"
JUDGE_MODEL = "gemini-2.5-flash-preview-05-20"
MAX_CONCURRENT_CALLS = 10

def find_latest_run_dir():
    """Finds the most recent benchmark run directory."""
    if not BENCHMARK_RUNS_DIR.exists():
        return None
    
    run_dirs = [d for d in BENCHMARK_RUNS_DIR.iterdir() if d.is_dir() and d.name.startswith('run_')]
    if not run_dirs:
        return None
        
    latest_run_dir = max(run_dirs, key=lambda d: d.stat().st_mtime)
    return latest_run_dir

def judge_single_item(row_tuple, run_dir):
    """
    Finds pre-generated images and calls the LLM judge.
    """
    index, row = row_tuple

    if not row.get('success', False) or not isinstance(row.get('before_images_path'), str) or not isinstance(row.get('after_images_path'), str):
        return index, {"status": "Skipped - Run not successful or image paths missing"}

    try:
        before_img_dir = run_dir / row['before_images_path']
        after_img_dir = run_dir / row['after_images_path']

        before_img_candidates = list(before_img_dir.glob("*.png"))
        if not before_img_candidates:
            return index, {"error": f"Before image not found in {before_img_dir}", "status": "Error"}
        
        after_img_candidates = list(after_img_dir.glob("*.png"))
        if not after_img_candidates:
            return index, {"error": f"After image not found in {after_img_dir}", "status": "Error"}

        before_img_path = before_img_candidates[0]
        generated_img_path = after_img_candidates[0]
        
        # ---> ADD XML EXTRACTION LOGIC <---
        before_ppt_path = run_dir / row['before_ppt_path']
        after_ppt_path = run_dir / row['output_pptx_path']
        
        # The 'modified_xml_files' column is a string representation of a list
        # Use json.loads to safely parse it back into a Python list
        try:
            modified_files_list = json.loads(row.get('modified_xml_files', '[]').replace("'", '"'))
        except json.JSONDecodeError:
            modified_files_list = []

        before_xml_content = {}
        after_xml_content = {}

        if modified_files_list:
            # print(f"\\nExtracting XML diffs for {row['id']}...")
            for xml_file in modified_files_list:
                before_xml = ppt_processor.extract_specific_xml_from_pptx(str(before_ppt_path), xml_file)
                after_xml = ppt_processor.extract_specific_xml_from_pptx(str(after_ppt_path), xml_file)
                if before_xml and after_xml:
                    before_xml_content[xml_file] = before_xml
                    after_xml_content[xml_file] = after_xml
        
        # ---> UPDATE THE JUDGE CALL <---
        judge_result = llm_handler.call_llm_judge(
            instruction=row['instruction'],
            before_img_path=str(before_img_path),
            after_img_path=str(generated_img_path),
            before_xml_dict=before_xml_content, # <-- Pass before XML
            after_xml_dict=after_xml_content,   # <-- Pass after XML
            model_id=JUDGE_MODEL
        )

        if 'error' in judge_result:
            return index, {"status": "Error", "error": judge_result.pop("error")}
        else:
            judge_result['status'] = 'Success'
        return index, judge_result

    except Exception as e:
        return index, {"error": f"An unexpected error occurred in judging framework: {e}", "status": "Error"}

def image_to_base64(image_path):
    """Converts an image file to a base64 string for embedding in HTML."""
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Warning: Could not encode image {image_path}: {e}")
        return None

def generate_html_report(df, html_output_path, run_dir):
    """Generates a self-contained HTML report with embedded images."""
    html_template = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Benchmark Evaluation Report</title>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 2em; background-color: #f8f9fa; color: #212529; }}
            h1 {{ color: #343a40; border-bottom: 2px solid #dee2e6; padding-bottom: 0.5em; }}
            h2 {{ color: #495057; margin: 0; }}
            .run-info {{ background: #e9ecef; padding: 1em; border-radius: 8px; margin-bottom: 2em; border: 1px solid #dee2e6; }}
            .result-card {{ border: 1px solid #dee2e6; border-radius: 8px; padding: 1.5em; margin-bottom: 1.5em; background: #ffffff; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
            .card-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 1em; }}
            .prompt {{ font-weight: bold; font-size: 1.1em; background-color: #f1f3f5; padding: 0.75em; border-radius: 4px; margin-bottom: 1em; }}
            .images {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1.5em; margin-bottom: 1em; }}
            .images div {{ text-align: center; }}
            .images img {{ max-width: 100%; height: auto; border: 1px solid #ced4da; border-radius: 4px; }}
            .scores {{ font-size: 0.9em; font-weight: bold; color: #343a40; padding: 0.5em; background-color: #f8f9fa; border-radius: 4px; border: 1px solid #dee2e6; display: flex; gap: 1em; white-space: nowrap; }}
            .success {{ color: #28a745; font-weight: bold; }}
            .failure {{ color: #dc3545; font-weight: bold; }}
            .error {{ font-family: "SF Mono", "Menlo", "Monaco", "Consolas", "Liberation Mono", "Courier New", monospace; color: #dc3545; background: #f8d7da; padding: 0.5em; border-radius: 4px; white-space: pre-wrap; word-wrap: break-word; }}
        </style>
    </head>
    <body>
        <h1>Benchmark Evaluation Report</h1>
        <div class="run-info"><strong>Run Directory:</strong> {run_dir_name}</div>
        {report_body}
    </body>
    </html>
    """
    
    report_body = []
    for _, row in df.iterrows():
        
        scores_html = ""
        if 'judge_status' in row and pd.notna(row['judge_status']):
            if row.get('judge_status') == 'Success':
                scores_html += "<div class='scores'>"
                # --- FIX: Changed keys to snake_case to match JSON from LLM Judge ---
                scores_html += f"<span>Instr: {row.get('judge_instruction_following', '?')}</span>"
                scores_html += f"<span>Text: {row.get('judge_text_quality', '?')}</span>"
                scores_html += f"<span>Image: {row.get('judge_image_quality', '?')}</span>"
                scores_html += f"<span>Layout: {row.get('judge_layout_quality', '?')}</span>"
                scores_html += f"<span>Color: {row.get('judge_color_quality', '?')}</span>"
                scores_html += "</div>"
            elif row.get('judge_status') in ['Error', 'Skipped', 'Framework Error']:
                error_msg = row.get('judge_error', 'Evaluation Skipped')
                scores_html += f"<div class='error' style='padding:0.5em; margin:0;'><strong>Judge Status: {row.get('judge_status')}</strong>: {error_msg}</div>"

        card_content = f"<div class='card-header'><h2>Prompt ID: {row['id']}</h2>{scores_html}</div>"
        card_content += f"<div class='prompt'>Instruction: {row['instruction']}</div>"

        if row['success']:
            card_content += "<p><strong>Status:</strong> <span class='success'>Success</span></p>"
            
            before_img_path = run_dir / row.get('before_images_path', '')
            after_img_path = run_dir / row.get('after_images_path', '')
            
            before_img_file = next(before_img_path.glob("*.png"), None) if before_img_path.exists() else None
            after_img_file = next(after_img_path.glob("*.png"), None) if after_img_path.exists() else None

            images_html = "<div class='images'>"
            if before_img_file:
                before_b64 = image_to_base64(before_img_file)
                images_html += f"<div><h3>Before</h3><img src='data:image/png;base64,{before_b64}' alt='Before image'></div>"
            
            if after_img_file:
                after_b64 = image_to_base64(after_img_file)
                images_html += f"<div><h3>After</h3><img src='data:image/png;base64,{after_b64}' alt='After image'></div>"
            else:
                 images_html += "<div><h3>After</h3><p>Image not generated.</p></div>"

            images_html += "</div>"
            card_content += images_html

        else:
            card_content += "<p><strong>Status:</strong> <span class='failure'>Failure</span></p>"
            if pd.notna(row.get('error_message')):
                card_content += f"<div class='error'><strong>Error Message:</strong> {row['error_message']}</div>"
        
        report_body.append(f"<div class='result-card'>{card_content}</div>")

    final_html = html_template.format(run_dir_name=run_dir.name, report_body="\n".join(report_body))
    
    with open(html_output_path, 'w', encoding='utf-8') as f:
        f.write(final_html)
    print(f"\nGenerated HTML report: {html_output_path}")

def evaluate_latest_run():
    """
    Finds the latest benchmark run, judges results, and saves an evaluated CSV and HTML report.
    """
    latest_run_dir = find_latest_run_dir()
    if not latest_run_dir:
        print("Error: No benchmark run directories found in 'benchmark_runs'.")
        print("Please run benchmark_runner.py first.")
        return

    print(f"Evaluating latest benchmark run: {latest_run_dir.name}")
    
    results_csv_path = latest_run_dir / "benchmark_results.csv"
    if not results_csv_path.exists():
        print(f"Error: Results file not found at {results_csv_path}.")
        return

    evaluated_csv_path = latest_run_dir / "benchmark_eval_scores.csv"
    html_report_path = latest_run_dir / "evaluation_report.html"

    df = pd.read_csv(results_csv_path).fillna('')
    
    # Filter to only judge successful runs
    successful_runs = df[df['success'] == True]
    print(f"Starting evaluation for {len(successful_runs)} successful prompts from: {results_csv_path.name}")
    print(f"Using {JUDGE_MODEL} as the judge with up to {MAX_CONCURRENT_CALLS} concurrent calls.")

    scores = [{} for _ in range(len(df))]
    
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_CALLS) as executor:
        tasks = [(row, latest_run_dir) for row in successful_runs.iterrows()]
        future_to_index = {executor.submit(judge_single_item, *task): task[0][0] for task in tasks}
        
        for future in tqdm(as_completed(future_to_index), total=len(tasks), desc="Judging Slides"):
            index = future_to_index[future]
            try:
                _, data = future.result()
                scores[index] = data
            except Exception as exc:
                scores[index] = {"error": str(exc), "status": "Framework Error"}

    scores_df = pd.DataFrame(scores)
    if not scores_df.empty:
        scores_df = scores_df.add_prefix('judge_')
        evaluated_df = df.join(scores_df)
    else:
        evaluated_df = df
    
    evaluated_df.to_csv(evaluated_csv_path, index=False)
    print(f"\nEvaluation complete. Full results with scores saved to {evaluated_csv_path}")

    generate_html_report(evaluated_df, html_report_path, latest_run_dir)

if __name__ == "__main__":
    evaluate_latest_run()