#!/usr/bin/env python3
"""
paddle-ocrapi.py
----------------
A robust utility script to perform OCR on PDF and image invoices using the
Paddle-OCR API (via AIStudio).

It uploads a file (or multiple files from a folder) to the API, polls for job
completion, downloads the structured layout parsing output, and saves the
extracted text with page separators, making it compatible with invoice extraction.

Requirements:
    pip install requests

Usage:
    python paddle-ocrapi.py -i invoices/
    python paddle-ocrapi.py -i invoices/sample.pdf -o extracted_text/
"""

import os
import sys
import json
import time
import argparse
from pathlib import Path
import requests

# API Settings
JOB_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
DEFAULT_MODEL = "PaddleOCR-VL-1.6"

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

def _load_env_token() -> str:
    # Try importing dotenv first, else fallback to reading manually
    try:
        from dotenv import load_dotenv
        load_dotenv(BASE_DIR / ".env")
    except ImportError:
        dotenv_path = BASE_DIR / ".env"
        if dotenv_path.exists():
            with open(dotenv_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" in line:
                        k, v = line.split("=", 1)
                        if k.strip() == "PADDLE_OCR_TOKEN":
                            return v.strip().strip('"').strip("'")
    return os.environ.get("PADDLE_OCR_TOKEN", "5e3cb5568aa2cac4e5bccb23888a9e1f061754cd")

DEFAULT_TOKEN = _load_env_token()

def log_status(msg: str) -> None:
    """Print status message with standard format."""
    print(f"[*] {msg}")

def log_success(msg: str) -> None:
    """Print success message."""
    print(f"[+] {msg}")

def log_error(msg: str) -> None:
    """Print error message."""
    print(f"[-] ERROR: {msg}", file=sys.stderr)

def submit_job(file_path: Path, token: str, model: str) -> str:
    """
    Submits a local file to the Paddle-OCR API.
    Returns the jobId if successful.
    """
    headers = {
        "Authorization": f"bearer {token}",
    }
    
    optional_payload = {
        "useDocOrientationClassify": False,
        "useDocUnwarping": False,
        "useChartRecognition": False,
    }

    log_status(f"Submitting job for: {file_path.name}")

    if not file_path.exists():
        raise FileNotFoundError(f"Local file not found: {file_path}")

    data = {
        "model": model,
        "optionalPayload": json.dumps(optional_payload)
    }

    with open(file_path, "rb") as f:
        files = {"file": f}
        response = requests.post(JOB_URL, headers=headers, data=data, files=files)

    if response.status_code != 200:
        raise requests.HTTPError(f"Submission failed (HTTP {response.status_code}): {response.text}")

    resp_json = response.json()
    if resp_json.get("code") != 0 or "data" not in resp_json:
        raise ValueError(f"API returned error status ({resp_json.get('code')}): {resp_json.get('msg', 'Unknown error')}")

    job_id = resp_json["data"]["jobId"]
    log_success(f"Job submitted successfully. Job ID: {job_id}")
    return job_id

def poll_job(job_id: str, token: str) -> str:
    """
    Polls the Paddle-OCR API until the job state is done or failed.
    Returns the result JSONL URL.
    """
    headers = {
        "Authorization": f"bearer {token}",
    }
    
    log_status("Polling for job status...")
    start_time = time.time()
    
    while True:
        try:
            response = requests.get(f"{JOB_URL}/{job_id}", headers=headers)
            if response.status_code != 200:
                log_error(f"Failed to fetch job status (HTTP {response.status_code}). Retrying...")
                time.sleep(5)
                continue
                
            resp_json = response.json()
            if resp_json.get("code") != 0 or "data" not in resp_json:
                raise ValueError(f"Status check failed ({resp_json.get('code')}): {resp_json.get('msg', 'Unknown error')}")
                
            job_data = resp_json["data"]
            state = job_data.get("state")
            elapsed = int(time.time() - start_time)
            
            if state == 'pending':
                log_status(f"Job is pending... (elapsed: {elapsed}s)")
            elif state == 'running':
                progress = job_data.get("extractProgress", {})
                total = progress.get("totalPages", "?")
                extracted = progress.get("extractedPages", "0")
                log_status(f"Job is running... Progress: {extracted}/{total} pages (elapsed: {elapsed}s)")
            elif state == 'done':
                progress = job_data.get("extractProgress", {})
                extracted = progress.get("extractedPages", "0")
                log_success(f"Job completed. Pages extracted: {extracted}. (elapsed: {elapsed}s)")
                
                result_url = job_data.get("resultUrl", {})
                json_url = result_url.get("jsonUrl")
                if not json_url:
                    raise ValueError("Job succeeded but JSON URL was not provided in result.")
                return json_url
            elif state == 'failed':
                error_msg = job_data.get("errorMsg", "Unknown failure reason")
                raise RuntimeError(f"Job failed: {error_msg}")
            else:
                log_status(f"Unknown job state '{state}'... (elapsed: {elapsed}s)")
                
        except Exception as e:
            if isinstance(e, (RuntimeError, ValueError)):
                raise e
            log_error(f"Error checking job status: {e}. Retrying in 5s...")
            
        time.sleep(5)

def download_and_process_result(
    jsonl_url: str,
    output_dir: Path,
    media_dir: Path,
    file_stem: str,
    save_media: bool
) -> Path:
    """
    Downloads the JSONL result, parses it, extracts text with page separators,
    and optionally downloads layout/markdown media.
    Returns the path to the saved text file.
    """
    log_status(f"Downloading result from: {jsonl_url}")
    response = requests.get(jsonl_url)
    response.raise_for_status()
    
    lines = response.text.strip().split('\n')
    extracted_pages = []
    
    os.makedirs(output_dir, exist_ok=True)
    
    # Media download directory for this specific file
    file_media_dir = media_dir / file_stem
    if save_media:
        os.makedirs(file_media_dir, exist_ok=True)
        
    page_num = 0
    for line_num, line in enumerate(lines, start=1):
        line = line.strip()
        if not line:
            continue
            
        try:
            line_data = json.loads(line)
        except json.JSONDecodeError:
            log_error(f"Failed to parse JSONL line {line_num}")
            continue
            
        result = line_data.get("result", {})
        layout_results = result.get("layoutParsingResults", [])
        
        for res in layout_results:
            markdown_data = res.get("markdown", {})
            page_text = markdown_data.get("text", "")
            extracted_pages.append(page_text)
            
            if save_media:
                # Save individual page markdown
                page_md_path = file_media_dir / f"page_{page_num + 1}.md"
                with open(page_md_path, "w", encoding="utf-8") as md_file:
                    md_file.write(page_text)
                
                # Save markdown images
                images = markdown_data.get("images", {})
                for img_path_str, img_url in images.items():
                    try:
                        # Clean up path/subfolders inside media_dir
                        target_img_path = file_media_dir / img_path_str
                        target_img_path.parent.mkdir(parents=True, exist_ok=True)
                        
                        img_bytes = requests.get(img_url).content
                        with open(target_img_path, "wb") as img_file:
                            img_file.write(img_bytes)
                        log_status(f"Saved layout image: {img_path_str}")
                    except Exception as e:
                        log_error(f"Failed to save layout image '{img_path_str}': {e}")
                        
                # Save output images
                output_images = res.get("outputImages", {})
                for img_name, img_url in output_images.items():
                    try:
                        target_img_path = file_media_dir / f"{img_name}_page_{page_num + 1}.jpg"
                        img_resp = requests.get(img_url)
                        if img_resp.status_code == 200:
                            with open(target_img_path, "wb") as f:
                                f.write(img_resp.content)
                            log_status(f"Saved output image: {target_img_path.name}")
                        else:
                            log_error(f"Failed to fetch image {img_name} (HTTP {img_resp.status_code})")
                    except Exception as e:
                        log_error(f"Failed to save output image '{img_name}': {e}")
            
            page_num += 1

    # Format extracted text with standard page separators
    formatted_text = ""
    for idx, page_text in enumerate(extracted_pages, start=1):
        formatted_text += f"--- Page {idx} ---\n{page_text}\n\n"
        
    output_text_path = output_dir / f"{file_stem}__paddleocr.txt"
    with open(output_text_path, "w", encoding="utf-8") as f:
        f.write(formatted_text.strip() + "\n")
        
    log_success(f"Extracted text saved to: {output_text_path}")
    return output_text_path

def process_file(
    file_path: Path,
    token: str,
    model: str,
    output_dir: Path,
    media_dir: Path,
    save_media: bool
) -> bool:
    """
    Submits, polls, and downloads OCR data for a single file.
    """
    try:
        job_id = submit_job(file_path, token, model)
        jsonl_url = poll_job(job_id, token)
        download_and_process_result(
            jsonl_url=jsonl_url,
            output_dir=output_dir,
            media_dir=media_dir,
            file_stem=file_path.stem,
            save_media=save_media
        )
        return True
    except Exception as e:
        log_error(f"Failed to process '{file_path.name}': {e}")
        return False

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Paddle-OCR API client for extracting document/invoice layout and text."
    )
    parser.add_argument(
        "-i", "--input",
        type=str,
        help="Path to an input file (PDF/image) or directory containing files. Defaults to the 'invoices' folder."
    )
    parser.add_argument(
        "-o", "--output-dir",
        type=str,
        default="extracted_text",
        help="Directory to save the formatted extracted text. (default: 'extracted_text')"
    )
    parser.add_argument(
        "-m", "--media-dir",
        type=str,
        default="output/paddleocr",
        help="Directory to save the layouts, markdown, and image assets. (default: 'output/paddleocr')"
    )
    parser.add_argument(
        "-t", "--token",
        type=str,
        default=DEFAULT_TOKEN,
        help="Paddle-OCR API Authorization Token."
    )
    parser.add_argument(
        "--model",
        type=str,
        default=DEFAULT_MODEL,
        help=f"Model to use (default: {DEFAULT_MODEL})."
    )
    parser.add_argument(
        "--no-media",
        action="store_true",
        help="Skip downloading and saving markdown pages and layout/output images."
    )

    args = parser.parse_args()

    # Determine input target
    input_path_str = args.input
    if not input_path_str:
        # Check standard default folders
        default_invoices = Path("invoices")
        if default_invoices.is_dir():
            input_path_str = str(default_invoices)
            log_status(f"No input specified. Defaulting to directory: {default_invoices}")
        else:
            log_error("No input specified and 'invoices' directory not found.")
            parser.print_help()
            sys.exit(1)

    input_path = Path(input_path_str)
    output_dir = Path(args.output_dir)
    media_dir = Path(args.media_dir)
    save_media = not args.no_media

    if not input_path.exists():
        log_error(f"Input path does not exist: {input_path}")
        sys.exit(1)

    # Gather files to process
    files_to_process = []
    supported_extensions = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}

    if input_path.is_dir():
        for p in sorted(input_path.iterdir()):
            if p.is_file() and p.suffix.lower() in supported_extensions:
                files_to_process.append(p)
        if not files_to_process:
            log_error(f"No supported files found in directory: {input_path}")
            sys.exit(1)
        log_status(f"Found {len(files_to_process)} file(s) in directory '{input_path}' to process.")
    else:
        if input_path.suffix.lower() not in supported_extensions:
            log_error(f"Unsupported file format: {input_path.suffix}")
            sys.exit(1)
        files_to_process.append(input_path)

    # Process all files
    success_count = 0
    for file_path in files_to_process:
        print("\n" + "=" * 60)
        log_status(f"Processing file: {file_path.name}")
        print("=" * 60)
        
        success = process_file(
            file_path=file_path,
            token=args.token,
            model=args.model,
            output_dir=output_dir,
            media_dir=media_dir,
            save_media=save_media
        )
        if success:
            success_count += 1

    print("\n" + "=" * 60)
    log_success(f"Processing summary: {success_count}/{len(files_to_process)} file(s) successfully processed.")
    print("=" * 60)

if __name__ == "__main__":
    main()