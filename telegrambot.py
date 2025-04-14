import requests
import os
import time
import json
import PyPDF2
import google.generativeai as genai
from bs4 import BeautifulSoup # For HTML parsing
import pandas as pd          # For Excel/CSV logging AND reading
import atexit                # To save log on exit
import io                    # To handle image bytes in memory
import mimetypes             # To guess image MIME type
import re # For Markdown escaping

# --- Configuration ---
# --- IMPORTANT: Replace placeholders with your actual credentials/paths ---
# Consider using environment variables for sensitive data in production
BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "") # <--- REPLACE or set Environment Variable
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "") # <--- REPLACE or set Environment Variable
BOT_USERNAME = "" # Replace if your bot has a different username
COMPANY_NAME = "" # Replace with your company name
# --- IMPORTANT: Use the correct path separator for your OS (\ for Windows, / for Linux/macOS) ---
COMPANY_INFO_FOLDER = "" # <--- REPLACE with the ACTUAL path to your folder
ADMIN_USER_ID =  # <--- REPLACE with YOUR Telegram User ID to receive admin alerts
EXCEL_LOG_FILE = "bot_conversation_log.xlsx" # Name for the conversation log file

# --- Basic Validation ---
if BOT_TOKEN == "YOUR_TELEGRAM_BOT_TOKEN" or not BOT_TOKEN:
    print("ERROR: Telegram Bot Token is not set. Please edit the script or set the TELEGRAM_BOT_TOKEN environment variable.")
    exit()
if GEMINI_API_KEY == "YOUR_GEMINI_API_KEY" or not GEMINI_API_KEY:
    print("ERROR: Gemini API Key is not set. Please edit the script or set the GEMINI_API_KEY environment variable.")
    exit()
if not os.path.isdir(COMPANY_INFO_FOLDER):
     print(f"ERROR: Company Info Folder '{COMPANY_INFO_FOLDER}' not found. Please check the path.")
     # Decide if you want to exit or continue without company data
     # exit() # Option to exit if folder is crucial
     print("Warning: Proceeding without loading company data.")


# Telegram API URL
TELEGRAM_API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"

# --- Initialize Gemini API ---
try:
    genai.configure(api_key=GEMINI_API_KEY)
    # Use a model compatible with Vision and strong text processing
    model = genai.GenerativeModel("gemini-1.5-pro")
    print("Gemini API configured successfully.")
except Exception as e:
    print(f"ERROR configuring Gemini API: {e}. Check your API key and network connection.")
    exit() # Exit if Gemini can't be configured

# --- Load Company Information ---
company_data = {}
intent_data = {} # Keep for potential simple keyword matching

def load_company_info():
    """
    Loads company information from PDF, JSON, HTML, CSV, and Excel files
    found in the COMPANY_INFO_FOLDER.
    Handles potential encoding issues in HTML/CSV.
    Converts tabular data (CSV/Excel) to JSON string format for better LLM processing.
    """
    global intent_data, company_data # Ensure we modify the global variables
    company_data = {} # Reset data on reload if called multiple times
    intent_data = {}

    print(f"Loading company information from: {COMPANY_INFO_FOLDER}")
    if not os.path.isdir(COMPANY_INFO_FOLDER):
        # This case might already be handled by the initial check, but good to have safety here too
        print(f"Error: Company info folder not found at {COMPANY_INFO_FOLDER} during load.")
        return # Stop loading if folder doesn't exist

    for filename in os.listdir(COMPANY_INFO_FOLDER):
        filepath = os.path.join(COMPANY_INFO_FOLDER, filename)
        if not os.path.isfile(filepath):
            continue # Skip directories or other non-files

        # --- PDF Processing ---
        if filename.lower().endswith(".pdf"):
            print(f"Processing PDF: {filename}")
            try:
                with open(filepath, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    text = ""
                    for page in pdf_reader.pages:
                        extracted = page.extract_text()
                        if extracted:
                             text += extracted + "\n" # Add newline between pages
                    if text.strip():
                        # Use filename without extension as key
                        company_data[os.path.splitext(filename)[0]] = text.strip()
                        print(f"  Successfully loaded content from {filename}")
                    else:
                        print(f"  Warning: No text extracted from {filename}")
            except Exception as e:
                print(f"  Error reading PDF {filename}: {e}")

        # --- JSON Processing ---
        elif filename.lower().endswith(".json"):
            print(f"Processing JSON: {filename}")
            try:
                # Attempt with UTF-8 first, common for JSON
                with open(filepath, 'r', encoding='utf-8') as json_file:
                    json_content = json.load(json_file)

                    # Convert content to string for Gemini context
                    if isinstance(json_content, (dict, list)):
                        json_string = json.dumps(json_content, indent=2) # Pretty print if complex
                    else:
                        json_string = str(json_content) # Simple conversion for primitives

                    company_data[os.path.splitext(filename)[0]] = json_string
                    print(f"  Loaded JSON content from {filename} as general information.")

                    # Optional: Check for specific 'intents' structure
                    if isinstance(json_content, dict) and 'intents' in json_content and isinstance(json_content['intents'], dict):
                        intent_data.update(json_content['intents'])
                        print(f"  Found and merged intents from {filename}.")

            except json.JSONDecodeError as e:
                print(f"  Error parsing JSON in {filename}: {e}")
            except UnicodeDecodeError:
                 print(f"  Error reading JSON {filename}: Not valid UTF-8. Skipping.")
            except Exception as e:
                print(f"  Error reading JSON file {filename}: {e}")

        # --- HTML Processing (with encoding fallback) ---
        elif filename.lower().endswith((".html", ".htm")):
            print(f"Processing HTML: {filename}")
            html_content = None
            # List of encodings to try, starting with the most common
            encodings_to_try = ['utf-8', 'cp1252', 'latin-1']
            for enc in encodings_to_try:
                try:
                    with open(filepath, 'r', encoding=enc) as html_file:
                        html_content = html_file.read()
                    print(f"  Successfully read HTML using encoding: {enc}")
                    break # Stop trying encodings if successful
                except UnicodeDecodeError:
                    print(f"  Failed to decode HTML with {enc}...")
                    continue # Try the next encoding in the list
                except Exception as e:
                    # Handle other potential errors during file read (e.g., permissions)
                    print(f"  Error reading HTML file {filename} (before parsing) with {enc}: {e}")
                    html_content = None # Ensure it's None if reading fails
                    break # Stop trying if a non-encoding error occurs

            if html_content: # Proceed only if reading was successful
                try:
                    # Use BeautifulSoup to parse the HTML content
                    # 'lxml' is generally faster, but 'html.parser' is built-in if lxml is not installed
                    soup = BeautifulSoup(html_content, 'lxml')
                    # Extract text content, separating elements with spaces and stripping whitespace
                    text = soup.get_text(separator=' ', strip=True)

                    if text.strip(): # Check if any text was actually extracted
                        company_data[os.path.splitext(filename)[0]] = text
                        print(f"  Successfully extracted text content from {filename}")
                    else:
                        print(f"  Warning: No text extracted from HTML {filename} after parsing (file might be script-heavy or empty).")
                except ImportError:
                    # Fallback or error if lxml is not installed
                    try:
                        print("  'lxml' parser not found, trying built-in 'html.parser'. (For better performance, 'pip install lxml')")
                        soup = BeautifulSoup(html_content, 'html.parser')
                        text = soup.get_text(separator=' ', strip=True)
                        if text.strip():
                            company_data[os.path.splitext(filename)[0]] = text
                            print(f"  Successfully extracted text content from {filename} using 'html.parser'")
                        else:
                             print(f"  Warning: No text extracted from HTML {filename} after parsing with 'html.parser'.")
                    except Exception as e_parse:
                         print(f"  Error parsing HTML content from {filename} even with html.parser: {e_parse}")

                except Exception as e_parse:
                    print(f"  Error parsing HTML content from {filename} with lxml: {e_parse}")
            else:
                 # This message appears if all encoding attempts failed during read
                 print(f"  Error: Could not read HTML file {filename} with any tried encodings ({', '.join(encodings_to_try)}). Skipping.")


        # --- CSV Processing ---
        elif filename.lower().endswith(".csv"):
            print(f"Processing CSV: {filename}")
            df = None # Initialize DataFrame variable
            encodings_to_try = ['utf-8', 'cp1252', 'latin-1'] # Common CSV encodings
            for enc in encodings_to_try:
                try:
                    # Read CSV into pandas DataFrame
                    df = pd.read_csv(filepath, encoding=enc)
                    print(f"  Successfully read CSV using encoding: {enc}")
                    break # Stop trying if read succeeds
                except UnicodeDecodeError:
                    print(f"  Failed to decode CSV with {enc}...")
                    df = None # Reset df if decoding fails
                    continue # Try next encoding
                except pd.errors.EmptyDataError:
                    print(f"  Warning: CSV file {filename} is empty.")
                    df = None
                    break # Stop trying if empty
                except Exception as e:
                    print(f"  Error reading CSV {filename} with {enc}: {e}")
                    df = None # Reset df on other errors
                    # Decide whether to break or continue trying other encodings on general errors
                    break # Often best to stop if it's not an encoding issue

            # Process DataFrame if it was loaded successfully and is not empty
            if df is not None and not df.empty:
                try:
                    # Convert dataframe to a JSON string (list of records format)
                    # This structure is generally easier for LLMs to understand than raw CSV text
                    # Handle potential date columns better during conversion
                    csv_json_string = df.to_json(orient='records', indent=2, date_format='iso')
                    company_data[os.path.splitext(filename)[0]] = f"Data from {filename}:\n{csv_json_string}"
                    print(f"  Successfully loaded and converted data from {filename}")
                except Exception as e_convert:
                     print(f"  Error converting CSV DataFrame to JSON for {filename}: {e_convert}")
                     # Optionally, store the raw string version as a fallback
                     # company_data[os.path.splitext(filename)[0]] = f"Data from {filename} (raw):\n{df.to_string()}"

            elif df is None and not pd.errors.EmptyDataError: # Only print error if it wasn't just an empty file
                print(f"  Error: Could not load CSV data from {filename} after trying encodings.")


        # --- Excel Processing ---
        elif filename.lower().endswith((".xlsx", ".xls")):
            print(f"Processing Excel: {filename}")
            excel_engine = None
            if filename.lower().endswith(".xlsx"):
                 excel_engine = 'openpyxl' # Use openpyxl for modern Excel files

            try:
                # Read all sheets into a dictionary of DataFrames {sheet_name: DataFrame}
                excel_data = pd.read_excel(filepath, sheet_name=None, engine=excel_engine)
                combined_excel_string = ""

                if excel_data: # Check if the dictionary of sheets is not empty
                    print(f"  Found sheets: {', '.join(excel_data.keys())}")
                    for sheet_name, df in excel_data.items():
                        if not df.empty:
                            try:
                                # Convert each sheet's DataFrame to JSON string
                                sheet_json_string = df.to_json(orient='records', indent=2, date_format='iso')
                                combined_excel_string += f"--- Sheet: {sheet_name} ---\n{sheet_json_string}\n\n"
                            except Exception as e_convert:
                                print(f"    Error converting sheet '{sheet_name}' to JSON in {filename}: {e_convert}")
                                # Fallback: Add sheet as string if JSON fails?
                                # combined_excel_string += f"--- Sheet: {sheet_name} (raw data) ---\n{df.to_string()}\n\n"
                        else:
                            # Optionally note empty sheets
                            combined_excel_string += f"--- Sheet: {sheet_name} ---\n(Sheet is empty)\n\n"

                    if combined_excel_string.strip():
                         company_data[os.path.splitext(filename)[0]] = f"Data from {filename}:\n{combined_excel_string.strip()}"
                         print(f"  Successfully loaded and converted data from {filename}")
                    else:
                         print(f"  Warning: All sheets processed in Excel file {filename} resulted in no data or were empty.")
                else:
                     print(f"  Warning: Could not read any sheets from Excel file {filename}. It might be empty or corrupted.")

            except ImportError:
                 # Provide specific instructions based on the file extension
                 if excel_engine == 'openpyxl':
                     print(f"  Error: Need 'openpyxl' library to read .xlsx files. Please install it: pip install openpyxl")
                 else: # Likely an older .xls file
                     print(f"  Error: Need a library to read .xls files (e.g., 'xlrd'). Please install one: pip install xlrd")
                 print(f"  Cannot read {filename}.")
            except Exception as e:
                print(f"  Error processing Excel file {filename}: {e}")

        # --- Ignoring other files ---
        else:
            # This will ignore files like .txt, .docx, .png, etc.
            # You could add handlers for more types if needed.
            print(f"Ignoring file: {filename} (Does not have a supported extension: .pdf, .json, .html, .htm, .csv, .xlsx, .xls)")

# --- Initial Load of Company Data ---
load_company_info() # Load data when the script starts
print("-" * 30)
print(f"Company data loaded from {len(company_data)} files.")
if company_data:
    print(f"  Data sources: {list(company_data.keys())}")
print(f"Intent data loaded with intents: {list(intent_data.keys()) if intent_data else 'No specific intents loaded'}")
print("-" * 30)


# --- Conversation Log ---
conversation_log = [] # Initialize empty list to store log entries

def add_log_entry(chat_id, user_id, user_input, bot_response, input_type="text"):
    """Adds an entry to the conversation log."""
    # Basic sanitization or length limiting could be added here if needed
    log_entry = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "chat_id": chat_id,
        "user_id": user_id,
        "input_type": input_type, # 'text' or 'image'
        "user_input": user_input, # Text message or image caption
        "bot_response": bot_response,
    }
    conversation_log.append(log_entry)
    # Optional: Print confirmation that log entry was added
    # print(f"Logged interaction for user {user_id} in chat {chat_id}")

def save_log_to_excel():
    """Saves the conversation log to an Excel file."""
    if not conversation_log:
        print("\nNo conversation data recorded, skipping log save.")
        return
    try:
        print(f"\nAttempting to save conversation log to {EXCEL_LOG_FILE}...")
        df = pd.DataFrame(conversation_log)
        # Use openpyxl engine for .xlsx format
        df.to_excel(EXCEL_LOG_FILE, index=False, engine='openpyxl')
        print(f"Successfully saved {len(conversation_log)} log entries to {EXCEL_LOG_FILE}.")
    except ImportError:
         print(f"Error saving log: 'openpyxl' library is required to write Excel files. Please install it: pip install openpyxl")
    except Exception as e:
        print(f"Error saving conversation log to Excel: {e}")
        # Potential fallback: Save as CSV
        try:
            csv_log_file = EXCEL_LOG_FILE.replace('.xlsx', '.csv')
            print(f"Attempting fallback save to CSV: {csv_log_file}")
            df.to_csv(csv_log_file, index=False)
            print(f"Successfully saved log as CSV: {csv_log_file}")
        except Exception as csv_e:
            print(f"Error saving log as CSV fallback: {csv_e}")


# Register the save function to be called automatically when the script exits
atexit.register(save_log_to_excel)

def escape_markdown_v2(text):
    """Escapes special characters for Telegram MarkdownV2."""
    escaped_text = re.sub(r'([_*\[\]()~`>#+\-=|{}.!])', r'\\\1', text)
    return escaped_text

# --- Telegram API Functions ---
def get_updates(offset=None):
    """Fetches new updates from Telegram API."""
    url = f"{TELEGRAM_API_URL}/getUpdates"
    # Increased timeout for potentially slow connections; allow only message updates
    params = {"timeout": 60, "offset": offset, "allowed_updates": ["message"]}
    try:
        # Request timeout should be slightly longer than the server timeout
        response = requests.get(url, params=params, timeout=70)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()
    except requests.exceptions.Timeout:
        print("Telegram getUpdates request timed out.")
        # Return structure consistent with successful call but empty result
        return {"ok": True, "result": []}
    except requests.exceptions.RequestException as e:
        print(f"Error fetching Telegram updates: {e}")
        # Return structure indicating failure
        return {"ok": False, "result": [], "error_message": str(e)}
    except json.JSONDecodeError as e:
        print(f"Error decoding Telegram API response: {e}")
        return {"ok": False, "result": [], "error_message": f"JSON Decode Error: {e}"}


def send_message(chat_id, text, reply_to=None, parse_mode="MarkdownV2"): # Added parse_mode parameter and default to MarkdownV2
    """Sends a text message to a specific chat via Telegram API."""
    url = f"{TELEGRAM_API_URL}/sendMessage"
    # Ensure text is string and reasonably sized (Telegram has limits)
    payload = {
        "chat_id": chat_id,
        "text": str(text), # Explicitly convert to string
        "parse_mode": parse_mode # Use specified or default parse mode
        }
    if reply_to:
        payload["reply_to_message_id"] = reply_to

    try:
        response = requests.post(url, data=payload, timeout=30) # Timeout for sending
        response.raise_for_status()
        print(f"Sent message to chat_id {chat_id}") # Confirmation log
        return True # Indicate success
    except requests.exceptions.RequestException as e:
        print(f"Error sending message to chat_id {chat_id}: {e}")
        # Log the error response text if available
        if response is not None:
             print(f"Error response: {response.text}")
        return False # Indicate failure


def send_message_to_admin(text):
    """Sends a formatted alert message to the predefined ADMIN_USER_ID."""
    if ADMIN_USER_ID and ADMIN_USER_ID != 123456789: # Check if ID is set and not default
        print(f"Sending alert to admin (ID: {ADMIN_USER_ID})...")
        # Format the message clearly identifying it as an admin alert
        admin_text = f"*🤖 Admin Alert ({COMPANY_NAME} Bot) 🤖*\n\n{text}"
        send_message(ADMIN_USER_ID, admin_text) # Keep Markdown for Admin messages
    else:
        print("ADMIN_USER_ID not set or is default placeholder. Cannot send admin message.")
        # Optionally, log the admin message to console instead
        print(f"Admin Alert (Not Sent): {text}")


def get_file_path(file_id):
    """Gets the file path for a given Telegram file_id using the getFile method."""
    url = f"{TELEGRAM_API_URL}/getFile"
    params = {"file_id": file_id}
    try:
        response = requests.get(url, params=params, timeout=20)
        response.raise_for_status()
        data = response.json()
        if data.get("ok"):
            # The path needed to construct the download URL
            return data["result"]["file_path"]
        else:
            print(f"Telegram API error getting file path for {file_id}: {data.get('description')}")
            return None
    except requests.exceptions.RequestException as e:
        print(f"Network error getting file path for file_id {file_id}: {e}")
        return None


def download_file(file_path):
    """Downloads a file from Telegram given its file_path obtained from get_file_path."""
    # Construct the full file download URL
    url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
    try:
        response = requests.get(url, timeout=45) # Allow more time for larger files
        response.raise_for_status()
        return response.content # Return file content as raw bytes
    except requests.exceptions.RequestException as e:
        print(f"Error downloading file from path {file_path}: {e}")
        return None


# --- Gemini Interaction Function ---
def get_gemini_response(user_message, chat_context="", image_data=None, image_mime_type=None):
    """
    Generates a response using the Gemini model.
    Uses loaded company data (PDF, JSON, HTML, CSV, Excel) as primary context.
    Can optionally process an image along with the user message.
    """
    print("-" * 30)
    log_prefix = f"[Gemini Request]"
    if image_data:
        print(f"{log_prefix} Processing image ({image_mime_type or 'unknown type'}) with query: '{user_message or '<no text query>'}'")
    else:
        print(f"{log_prefix} Processing text message: {user_message}")

    # 1. Simple Keyword Intent Matching (Optional Pre-filter)
    # Keep this very basic or remove if Gemini should handle all interpretation
    matched_intent_response = None
    if not image_data and intent_data and user_message: # Only check for text messages if intents exist
        for intent_keyword, response_data in intent_data.items():
            if intent_keyword.lower() in user_message.lower():
                print(f"{log_prefix} Intent Matched (Keyword): '{intent_keyword}'")
                if isinstance(response_data, dict) and "response" in response_data:
                    matched_intent_response = response_data["response"]
                    break
                elif isinstance(response_data, str): # Allow simple string responses
                    matched_intent_response = response_data
                    break
                else:
                    print(f"{log_prefix} Warning: Malformed response data for intent '{intent_keyword}'")
                    # Don't break, maybe another keyword matches later

    if matched_intent_response:
         print(f"{log_prefix} Using predefined response for matched intent.")
         return matched_intent_response

    # 2. Gemini-Powered Response Generation using Company Context + Query (+ Image)
    print(f"{log_prefix} No simple intent matched or check skipped. Querying Gemini model...")

    # Combine all loaded company data into a single context string
    # Use filenames as separators for clarity
    company_info_context = "\n\n".join(
        [f"--- Information from: {filename} ---\n{data}" for filename, data in company_data.items()]
    )
    if not company_info_context:
        company_info_context = "(No company-specific information was loaded)"

    # Construct the main prompt for Gemini
    # This prompt guides the AI on its role, data sources, and expected behavior
    full_context_prompt = f"""
**Your Role:** You are {BOT_USERNAME}, an AI assistant for {COMPANY_NAME}.

**Your Goal:** Provide helpful, accurate, and professional assistance to users.

**Available Information:**
You have access to the following information about {COMPANY_NAME}. Prioritize using this information to answer questions. The information comes from various documents (PDFs, web pages, spreadsheets, etc.).

{company_info_context}

**Chat History (Recent messages, if provided):**
{chat_context or "(No recent chat history provided)"}

**Instructions for Responding:**
1.  **Analyze the Request:** Understand if the user is asking a question, providing an image for analysis, or making a general statement.
2.  **Use Company Info First:** If the query relates to {COMPANY_NAME}, consult the "Available Information" above *first*. Base your answer primarily on this data.
3.  **Image Analysis:** If an image is provided, describe it or answer questions about it. Relate it to {COMPANY_NAME} if possible based on the "Available Information" or the user's query.
4.  **General Knowledge:** If the query is unrelated to {COMPANY_NAME} and the provided info, you *may* use your general knowledge for common questions, but keep it brief and always identify yourself as an AI assistant for {COMPANY_NAME}. Do *not* answer highly speculative, personal, or inappropriate questions.
5.  **Acknowledge Limitations:** If you cannot answer using the provided information or general knowledge, clearly state that. For example: "I don't have specific details on that in my current information for {COMPANY_NAME}." or "Based on the information I have, I cannot answer that question."
6.  **Do Not Invent:** Never make up information about {COMPANY_NAME}, its products, services, or policies.
7.  **Be Professional:** Maintain a polite, helpful, and professional tone.
8.  **Conciseness:** Keep answers clear and to the point, unless the user asks for more detail.
9.  **Offer Help:** If you're unsure or cannot fully answer, you can offer to pass the query to a human representative at {COMPANY_NAME}. (e.g., "Would you like me to forward your question to our support team?")
10. **multilangual:** Keep answers in the language in which language you recieve message.
"""

    # Prepare the content parts for the Gemini API call
    prompt_parts = [full_context_prompt] # Start with the main instruction prompt

    # Add image data if present
    if image_data and image_mime_type:
        prompt_parts.append("\n**Image Provided for Analysis:**")
        prompt_parts.append({
            "mime_type": image_mime_type,
            "data": image_data # Raw image bytes
        })
        # Add the user's text query related to the image (if any)
        if user_message:
            prompt_parts.append(f"\n**User Query Regarding Image:** {user_message}")
        else:
            # If no text came with the image, ask Gemini for a description
            prompt_parts.append("\n**User Query Regarding Image:** Please describe this image and explain its relevance to {COMPANY_NAME} if applicable.")

    elif user_message: # If no image, just add the user's text message
         prompt_parts.append(f"\n\n**User's Current Message:** {user_message}")
    else:
         # Handle cases where there's somehow no message and no image (should be rare)
         print(f"{log_prefix} Error: No user message or image data provided to Gemini.")
         return "I seem to have received an empty request. Could you please try sending your message or image again?"

    # Add the final instruction for the AI's response format
    prompt_parts.append(f"\n\n**Your Response ({BOT_USERNAME}):**")

    # print(f"DEBUG: Sending prompt parts to Gemini: {[type(p) if not isinstance(p, str) else 'str' for p in prompt_parts]}") # Debug types
    # print(f"DEBUG: Prompt Text Context Snippet:\n{full_context_prompt[:500]}...") # Debug text part

    try:
        # Generate content using the model
        # Consider adding safety_settings if needed:
        # safety_settings=[
        #     {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
        #     # Add more categories as needed
        # ]
        response = model.generate_content(prompt_parts) # Send the list of parts

        # Extract the text response
        ai_response = response.text.strip()

        # --- Post-processing and Safety Check ---
        # Basic check if Gemini gave a generic refusal or empty response
        uncertain_phrases = [
            "i cannot fulfill that request", "i don't have enough information",
            "unable to provide assistance", "cannot answer", "i am sorry",
            "my knowledge cutoff", "as a large language model" # Phrases indicating it might be off-topic or refusing
            ]
        # Check for empty response or overly generic refusals
        is_uncertain = not ai_response or any(phrase in ai_response.lower() for phrase in uncertain_phrases)

        # Also check prompt feedback for blocked content, although response.text usually handles this
        if hasattr(response, 'prompt_feedback') and response.prompt_feedback.block_reason:
             print(f"{log_prefix} Gemini response blocked. Reason: {response.prompt_feedback.block_reason}")
             # Send generic refusal to user and notify admin
             admin_help_message = f"User query caused content block.\nQuery: '{user_message or '<Image Only>'}'\nBlock Reason: {response.prompt_feedback.block_reason}"
             send_message_to_admin(admin_help_message)
             return f"I'm sorry, I cannot process that request due to safety guidelines. Please ask something else related to {COMPANY_NAME}."

        if is_uncertain:
            print(f"{log_prefix} Gemini response seems uncertain or generic: '{ai_response[:100]}...'")
            # Optionally notify admin if Gemini seems unable to answer reasonably
            # admin_help_message = f"Gemini response seemed uncertain for query: '{user_message or '<Image Only>'}'\nResponse: {ai_response}"
            # send_message_to_admin(admin_help_message)

            # Provide a more helpful refusal to the user
            return f"I looked through the available information for {COMPANY_NAME}, but I couldn't find a specific answer to your query. Could you perhaps rephrase it, or would you like me to ask a human colleague for help?"

        # If response seems valid
        print(f"{log_prefix} Gemini Response: {ai_response[:200]}...") # Log snippet
        return ai_response

    except Exception as e:
        # Catch-all for errors during API call or response processing
        error_message = f"Error during Gemini API call for query: '{user_message or '<Image Only>'}'\nError: {e}"
        print(f"{log_prefix} {error_message}")
        # Check if the error might be due to the response object itself (e.g., invalid structure)
        if 'response' in locals() and hasattr(response, 'prompt_feedback'):
             print(f"Gemini Prompt Feedback (if available): {response.prompt_feedback}")

        send_message_to_admin(error_message) # Notify admin of the failure
        # Provide a generic error message to the user
        return "Sorry, I encountered an unexpected issue while trying to process your request. Please try again in a moment. If the problem persists, please contact our support team directly."


# --- Message Processing Function ---
def process_messages():
    """Main loop to get updates from Telegram and process messages."""
    last_update_id = None
    print(f"🚀 Starting {COMPANY_NAME} Bot ({BOT_USERNAME})... Listening for messages.")
    print(f"   Admin User ID for alerts: {ADMIN_USER_ID if (ADMIN_USER_ID and ADMIN_USER_ID != 123456789) else 'Not Set'}")
    print(f"   Logging conversations to: {EXCEL_LOG_FILE}")
    print(f"   Company info folder: {COMPANY_INFO_FOLDER}")
    print("-" * 30)
    print("Bot is running. Press Ctrl+C to stop.")

    while True:
        try: # Wrap the main loop in a try-except to catch unexpected interruptions
            updates = get_updates(last_update_id)

            if updates and updates.get("ok"):
                if not updates.get("result"): # No new messages
                    # Optional: print a heartbeat message if needed for monitoring
                    # print(".", end="", flush=True)
                    time.sleep(1) # Short sleep if no updates
                    continue

                # Process each update received
                for update in updates["result"]:
                    current_update_id = update["update_id"]
                    # Immediately update last_update_id to prevent reprocessing this update
                    last_update_id = current_update_id + 1

                    if "message" not in update:
                        # Skip updates that aren't messages (e.g., channel post edits)
                        continue

                    message = update["message"]
                    message_id = message["message_id"] # ID of the user's message
                    chat = message["chat"]
                    user = message["from"]
                    chat_id = chat["id"]
                    user_id = user["id"]
                    chat_type = chat["type"] # "private", "group", "supergroup", "channel"
                    user_name = user.get('first_name', 'Unknown User') # Get user's first name

                    # Initialize variables for this message
                    message_text = ""      # Text content or caption
                    image_data = None      # Raw image bytes
                    image_mime_type = None # e.g., 'image/jpeg'
                    input_type = "unknown" # Track if it's text, image, or other
                    should_respond = False # Flag to determine if bot should process/reply


                    # --- 1. Check for Text Message ---
                    if "text" in message:
                        message_text = message["text"].strip()
                        input_type = "text"
                        print(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Received Text | Chat: {chat_id} ({chat_type}) | User: {user_name} ({user_id}) | Msg: '{message_text[:100]}{'...' if len(message_text)>100 else ''}'")

                        # Determine if the bot should respond to this text message
                        if chat_type == "private":
                            should_respond = True # Always respond in private chats
                        elif chat_type in ["group", "supergroup"]:
                            # Respond in groups only if mentioned by username
                            if (f"@{BOT_USERNAME}".lower() in message_text.lower()):
                                print(f"  Bot mentioned in group chat.")
                                should_respond = True
                                # Clean the mention from the message for clearer context to Gemini
                                message_text = message_text.replace(f"@{BOT_USERNAME}", "").strip()
                            # Add checks for commands if needed:
                            # elif message_text.startswith('/'):
                            #    handle_command(message_text, chat_id, message_id) # Separate function for commands

                    # --- 2. Check for Photo Message ---
                    elif "photo" in message:
                        input_type = "image"
                        # Caption is the text sent with the image
                        message_text = message.get("caption", "").strip() # Use caption as query text
                        print(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Received Photo | Chat: {chat_id} ({chat_type}) | User: {user_name} ({user_id}) | Caption: '{message_text[:100]}{'...' if len(message_text)>100 else ''}'")

                        # Find the largest photo available (last in the 'photo' list)
                        photo_info = message['photo'][-1]
                        file_id = photo_info['file_id']
                        print(f"  Image File ID: {file_id}, Dimensions: {photo_info['width']}x{photo_info['height']}")

                        # Download the image file
                        file_path = get_file_path(file_id)
                        if file_path:
                            print(f"  Downloading image from path: {file_path}")
                            image_bytes = download_file(file_path)
                            if image_bytes:
                                image_data = image_bytes # Store raw bytes
                                # Guess MIME type from file path extension (e.g., .jpg -> image/jpeg)
                                guessed_type, _ = mimetypes.guess_type(file_path)
                                image_mime_type = guessed_type or 'image/jpeg' # Default if guess fails
                                print(f"  Downloaded {len(image_data)} bytes. Detected MIME type: {image_mime_type}")

                                # Determine if bot should respond to this image message
                                if chat_type == "private":
                                    should_respond = True # Respond to images in private chat
                                elif chat_type in ["group", "supergroup"]:
                                     # Respond in groups if mentioned in the caption
                                    if message_text and (f"@{BOT_USERNAME}".lower() in message_text.lower()):
                                        print(f"  Bot mentioned in image caption in group chat.")
                                        should_respond = True
                                        # Clean mention from caption
                                        message_text = message_text.replace(f"@{BOT_USERNAME}", "").strip()
                                    # Add logic here if you want the bot to respond to *all* images in groups (potentially noisy)
                                    # else:
                                    #    should_respond = True # Example: respond to all images

                            else:
                                print("  Failed to download image content.")
                                send_message(chat_id, "Sorry, I couldn't download the image you sent. Please try again.", message_id)
                        else:
                            print("  Failed to get file path for the image from Telegram.")
                            send_message(chat_id, "Sorry, there was a problem accessing the image file information.", message_id)

                    # --- 3. Handle Other Message Types (Optional) ---
                    # Example: Document, Audio, etc. - Currently ignored.
                    else:
                        message_type = next((key for key in message if key not in ['message_id', 'from', 'chat', 'date']), 'unknown')
                        print(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Received Unsupported Type | Chat: {chat_id} ({chat_type}) | User: {user_name} ({user_id}) | Type: {message_type}")
                        # Decide if you want to notify the user
                        # send_message(chat_id, f"Sorry, I can only process text messages and photos right now.", message_id)


                    # --- Generate and Send Response if needed ---
                    if should_respond and (message_text or image_data): # Need either text or image to proceed
                        try:
                            # Indicate bot is thinking (optional)
                            # requests.post(f"{TELEGRAM_API_URL}/sendChatAction", data={"chat_id": chat_id, "action": "typing"})

                            # Fetch context (e.g., chat history - currently disabled for simplicity)
                            chat_context = "" # Add get_chat_history(chat_id) if implemented and needed

                            # Call the Gemini function
                            response_text = get_gemini_response(
                                user_message=message_text,
                                chat_context=chat_context,
                                image_data=image_data,
                                image_mime_type=image_mime_type
                            )

                            # Escape MarkdownV2 characters in the response
                            escaped_response_text = escape_markdown_v2(response_text)

                            # Send the response back to the user with MarkdownV2 enabled
                            if escaped_response_text:
                                message_sent = send_message(chat_id, escaped_response_text, message_id, parse_mode="MarkdownV2") # Reply to original message, using MarkdownV2
                                # Log the interaction IF message was sent successfully
                                if message_sent:
                                    add_log_entry(chat_id, user_id, message_text or "[Image without caption]", response_text, input_type) # Log original response (un-escaped)
                            else:
                                # Handle case where Gemini function returns None or empty string (should be rare now)
                                print("Warning: Gemini function returned an empty response.")
                                error_fallback = "Sorry, I couldn't generate a response for that request."
                                send_message(chat_id, error_fallback, message_id)
                                add_log_entry(chat_id, user_id, message_text or "[Image without caption]", "[Error: Empty response from Gemini]", input_type)

                        except Exception as e_process:
                             # Catch unexpected errors during response generation or sending
                             print(f"!!! CRITICAL Error processing update {current_update_id} for chat {chat_id}: {e_process}")
                             # Send detailed error to admin
                             send_message_to_admin(f"CRITICAL Error processing update for chat {chat_id}:\nUser: {user_name} ({user_id})\nType: {input_type}\nError: {e_process}")
                             # Try to notify the user generically
                             error_fallback = "An unexpected error occurred while processing your request. The admin has been notified. Please try again later."
                             send_message(chat_id, error_fallback, message_id)
                             # Log the error interaction
                             add_log_entry(chat_id, user_id, message_text or "[Image Error]", f"[CRITICAL Error: {e_process}]", input_type)

                    elif not should_respond and input_type != "unknown":
                         # Log ignored messages if needed for debugging, but can be noisy
                         # print(f"  Ignoring {input_type} message (not private, bot not mentioned, or no content).")
                         pass


            elif updates is None:
                # Handle case where get_updates might return None (custom handling)
                print("Error fetching updates (get_updates returned None). Retrying after delay...")
                time.sleep(10) # Wait longer before retrying network issues

            elif not updates.get("ok"):
                # Handle explicit errors returned by Telegram API in get_updates
                error_code = updates.get("error_code")
                description = updates.get("description", "No description")
                print(f"Telegram API Error in getUpdates: Code {error_code} - {description}")
                if error_code == 401: # Unauthorized
                    print("CRITICAL: Telegram Bot Token is invalid or revoked. Stopping bot.")
                    break # Exit the loop
                elif error_code == 409: # Conflict
                    print("WARNING: Conflict detected (Error 409). Another bot instance might be running with the same token.")
                    print("Waiting 60 seconds before retrying...")
                    time.sleep(60)
                else:
                    # Wait before retrying other API errors
                    time.sleep(15)

            # Small delay between polling loops to avoid hammering the API
            time.sleep(1)

        except KeyboardInterrupt:
            print("\nCtrl+C detected. Stopping bot...")
            break # Exit the while loop cleanly
        except Exception as e_main_loop:
            # Catch any other unexpected errors in the main loop
            print(f"\n!!! UNHANDLED EXCEPTION IN MAIN LOOP: {e_main_loop}")
            send_message_to_admin(f"UNHANDLED EXCEPTION in bot main loop:\n{e_main_loop}")
            print("Restarting loop after 15 seconds...")
            time.sleep(15)


# --- Start Bot Execution ---
if __name__ == "__main__":
    # load_company_info() # Ensure data is loaded (already called after function definition)
    process_messages() # Start the main message processing loop
    print("Bot has stopped.")