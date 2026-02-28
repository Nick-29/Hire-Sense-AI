#!/usr/bin/env python3
"""
Twilio Recruitment Auto-Caller – Calls shortlisted candidates from Excel
Reads Shortlisted.xlsx, calls each candidate, updates Excel with call status and IVR responses.
"""

import os
import time
import threading
import logging
import urllib.parse
from datetime import datetime
import pandas as pd
from twilio.rest import Client
from flask import Flask, request
from pyngrok import ngrok
from dotenv import load_dotenv

load_dotenv()

# ============================================
# CONFIGURATION FROM ENVIRONMENT
# ============================================
TWILIO_ACCOUNT_SID = os.getenv("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.getenv("TWILIO_AUTH_TOKEN")
TWILIO_PHONE_NUMBER = os.getenv("TWILIO_PHONE_NUMBER")
FLOW_SID = os.getenv("FLOW_SID")
NGROK_AUTH_TOKEN = os.getenv("NGROK_AUTH_TOKEN")

EXCEL_FILE = os.getenv("SHORTLISTED_FILE")
OUTPUT_FILE = os.getenv("CALLS_OUTPUT_FILE", "Shortlisted_called.xlsx")

CALL_DELAY = int(os.getenv("CALL_DELAY", 5))
MAX_RANK_TO_CALL = int(os.getenv("MAX_RANK_TO_CALL", 10))
MAX_WAIT_TIME = int(os.getenv("MAX_WAIT_TIME", 120))
RESPONSE_CHECK_INTERVAL = int(os.getenv("RESPONSE_CHECK_INTERVAL", 10))
NGROK_PORT = int(os.getenv("NGROK_PORT", 5000))

# ============================================
# LOGGING SETUP
# ============================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('recruitment_calls.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================
# NGROK SETUP
# ============================================
def start_ngrok():
    try:
        if NGROK_AUTH_TOKEN:
            ngrok.set_auth_token(NGROK_AUTH_TOKEN)
            logger.info("ngrok auth token configured")
        tunnel = ngrok.connect(NGROK_PORT, "http")
        public_url = tunnel.public_url
        logger.info(f"ngrok tunnel started: {public_url}")
        return public_url
    except Exception as e:
        logger.error(f"Failed to start ngrok: {e}")
        return f"http://localhost:{NGROK_PORT}"

# ============================================
# FLASK WEBHOOK
# ============================================
app = Flask(__name__)

# Global state
df = None
call_data_store = {}
total_candidates_to_call = 0
responses_received = 0
responses_lock = threading.Lock()
df_lock = threading.Lock()
shutdown_event = threading.Event()

@app.route('/ivr-response', methods=['POST'])
def ivr_response():
    global df, responses_received
    raw_data = request.get_data(as_text=True).strip()
    logger.info(f"Raw webhook data: {repr(raw_data)}")

    call_sid = None
    digit = None
    name = None
    row_idx = None

    if raw_data.startswith('body='):
        encoded_body = raw_data[5:]
        decoded_body = urllib.parse.unquote(encoded_body)
        parts = decoded_body.split('\\n') if '\\n' in decoded_body else decoded_body.split('\n')
        for part in parts:
            part = part.strip()
            if '=' in part:
                key, value = part.split('=', 1)
                key = key.strip()
                value = value.strip()
                if key == 'CallSid':
                    call_sid = value
                elif key == 'Digits':
                    digit = value
                elif key == 'name':
                    name = value
                elif key == 'row_idx':
                    row_idx = value

    logger.info(f"Parsed - Call: {call_sid}, Digit: {digit}, Name: {name}, Row: {row_idx}")

    matched = False
    matched_row_idx = None
    matched_name = None

    if row_idx is not None:
        try:
            row_idx_int = int(row_idx)
            with df_lock:
                if df is not None and row_idx_int in df.index:
                    matched_row_idx = row_idx_int
                    matched_name = df.at[row_idx_int, 'name']
                    matched = True
                    logger.info(f"Matched by original row index: {matched_name}")
        except (ValueError, TypeError):
            pass

    if not matched and name and df is not None:
        with df_lock:
            name_matches = df[df['name'].str.lower() == name.lower()]
            if len(name_matches) == 1:
                matched_row_idx = name_matches.index[0]
                matched_name = name
                matched = True
                logger.info(f"Matched by name: {matched_name}")

    if not matched and call_sid:
        for exec_sid, data in call_data_store.items():
            if exec_sid == call_sid:
                matched_row_idx = data['row_idx']
                matched_name = data['name']
                matched = True
                logger.info(f"Matched by call SID: {matched_name}")
                break

    if not matched:
        logger.warning("Could not match call to any candidate.")
        return "", 200

    if digit == '1':
        response_text = 'Interested'
        logger.info(f"Candidate {matched_name} is INTERESTED")
    elif digit == '2':
        response_text = 'Not Interested'
        logger.info(f"Candidate {matched_name} is NOT INTERESTED")
    else:
        response_text = f'Invalid Input: {digit}'
        logger.warning(f"Invalid key pressed: {digit}")

    with df_lock:
        if df is not None and matched_row_idx is not None and matched_row_idx in df.index:
            df.at[matched_row_idx, 'ivr_response'] = response_text
            df.at[matched_row_idx, 'call_status'] = 'Completed'
            df.at[matched_row_idx, 'call_sid'] = call_sid if call_sid else 'unknown'
            df.at[matched_row_idx, 'call_timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            df.to_excel(OUTPUT_FILE, index=False)
            logger.info(f"Excel updated for {matched_name}")

    with responses_lock:
        responses_received += 1
        logger.info(f"Progress: {responses_received}/{total_candidates_to_call} responses")

    return "", 200

@app.route('/status-callback', methods=['POST'])
def status_callback():
    raw_data = request.get_data(as_text=True).strip()
    call_sid = None
    call_status = None

    if raw_data.startswith('body='):
        encoded_body = raw_data[5:]
        decoded_body = urllib.parse.unquote(encoded_body)
        parts = decoded_body.split('\\n') if '\\n' in decoded_body else decoded_body.split('\n')
        for part in parts:
            part = part.strip()
            if '=' in part:
                key, value = part.split('=', 1)
                if key == 'CallSid':
                    call_sid = value
                elif key == 'CallStatus':
                    call_status = value

    logger.info(f"Status Callback - Call: {call_sid}, Status: {call_status}")

    if call_status in ['no-answer', 'busy', 'failed', 'canceled']:
        for exec_sid, data in call_data_store.items():
            if exec_sid == call_sid:
                row_idx = data['row_idx']
                name = data['name']
                with df_lock:
                    if df is not None and row_idx in df.index:
                        df.at[row_idx, 'call_status'] = call_status.title()
                        df.at[row_idx, 'ivr_response'] = 'No Response'
                        df.at[row_idx, 'call_timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        df.to_excel(OUTPUT_FILE, index=False)
                        logger.info(f"Call {call_sid} {call_status} - updated for {name}")
                with responses_lock:
                    responses_received += 1
                    logger.info(f"Progress: {responses_received}/{total_candidates_to_call} responses")
                break
    return "", 200

def run_flask():
    app.run(host='0.0.0.0', port=NGROK_PORT, debug=False, use_reloader=False)

# ============================================
# EXCEL HANDLING
# ============================================
def load_and_filter_top_ranked(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    df = df.rename(columns={
        'Candidate Name': 'name',
        'Email': 'email',
        'Mobile': 'phone',
        'Rank': 'rank'
    })
    required = ['name', 'phone', 'email', 'rank']
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' missing in Excel file")

    df['rank'] = pd.to_numeric(df['rank'], errors='coerce')
    df = df.dropna(subset=['rank'])
    df = df.sort_values(by='rank', ascending=True).reset_index(drop=True)

    df_filtered = df[df['rank'] <= MAX_RANK_TO_CALL].copy()
    for col in ['call_status', 'ivr_response', 'call_sid', 'execution_sid', 'call_timestamp']:
        if col not in df_filtered.columns:
            df_filtered[col] = None

    logger.info(f"Loaded {len(df)} shortlisted candidates. To call: {len(df_filtered)}")
    return df, df_filtered

def format_phone_to_e164(phone):
    phone = str(phone).strip()
    digits = ''.join(filter(str.isdigit, phone))
    if not digits:
        return None
    if phone.startswith('+'):
        return phone
    else:
        if len(digits) == 10:
            return f"+91{digits}"
        elif len(digits) == 12 and digits.startswith('91'):
            return f"+{digits}"
        else:
            return f"+{digits}"

def initiate_call(name, phone, original_idx, public_url):
    client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
    to_number = format_phone_to_e164(phone)
    if not to_number:
        logger.error(f"Invalid phone number: {phone}")
        return None

    try:
        execution = client.studio.v2.flows(FLOW_SID).executions.create(
            to=to_number,
            from_=TWILIO_PHONE_NUMBER,
            parameters={
                'name': name,
                'row_idx': original_idx,
                'callback_url': f"{public_url}/ivr-response"
            }
        )
        logger.info(f"Call initiated to {name} ({to_number}) - Execution SID: {execution.sid}")
        call_data_store[execution.sid] = {'row_idx': original_idx, 'name': name, 'phone': to_number}
        return execution.sid
    except Exception as e:
        logger.error(f"Failed to call {name}: {str(e)}")
        return None

def wait_for_responses():
    global responses_received, total_candidates_to_call
    start_time = time.time()
    while not shutdown_event.is_set():
        time.sleep(RESPONSE_CHECK_INTERVAL)
        with responses_lock:
            current = responses_received
        if current >= total_candidates_to_call and total_candidates_to_call > 0:
            logger.info(f"All {total_candidates_to_call} responses received.")
            shutdown_event.set()
            break
        if time.time() - start_time > MAX_WAIT_TIME:
            logger.info(f"Timeout. Received {current}/{total_candidates_to_call} responses.")
            shutdown_event.set()
            break

# ============================================
# MAIN
# ============================================
def main():
    global df, total_candidates_to_call
    print("\n========================================")
    print("TWILIO RECRUITMENT AUTO-CALLER")
    print("========================================\n")

    public_url = start_ngrok()
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    time.sleep(2)
    logger.info(f"Flask server started on port {NGROK_PORT}")
    logger.info(f"IVR endpoint: {public_url}/ivr-response")

    try:
        df, candidates_to_call = load_and_filter_top_ranked(EXCEL_FILE)
        total_candidates_to_call = len(candidates_to_call)
    except Exception as e:
        logger.error(f"Error loading Excel: {e}")
        ngrok.kill()
        return

    if total_candidates_to_call == 0:
        logger.warning("No candidates to call.")
        ngrok.kill()
        return

    logger.info(f"Calling {total_candidates_to_call} candidates...")

    for original_idx, row in candidates_to_call.iterrows():
        name = row['name']
        phone = row['phone']
        rank = int(row['rank']) if pd.notna(row['rank']) else 'N/A'

        if pd.isna(phone) or not str(phone).strip():
            logger.warning(f"Skipping {name} (Rank {rank}): no phone")
            continue

        logger.info(f"Calling Rank #{rank}: {name}")
        execution_sid = initiate_call(name, phone, original_idx, public_url)

        with df_lock:
            df.at[original_idx, 'execution_sid'] = execution_sid
            df.at[original_idx, 'call_status'] = 'Initiated' if execution_sid else 'Failed'
            df.at[original_idx, 'call_timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            df.to_excel(OUTPUT_FILE, index=False)

        if original_idx != candidates_to_call.index[-1]:
            time.sleep(CALL_DELAY)

    logger.info("Waiting for responses...")
    wait_for_responses()
    ngrok.kill()
    logger.info("Done.")

if __name__ == "__main__":
    main()