import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pymupdf  # PyMuPDF
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import re
import os
import threading
import time
import queue
import csv
from datetime import datetime

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None

# ------------------ CONFIG ------------------
DEBUG = 0  # 1 = only process first page of each PDF
ASK_TIMETABLE_EVERY_RUN = True  # if True, asks once per unique subject set per run

# If your tesseract isn't on PATH, set it here (Windows example):
if pytesseract:
    pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# ---------------- SUBJECT MAP (extend as needed) ----------------
SUBJECT_CODE_MAP = {
    "ENGAG": "English A",
    "ENGBG": "English B",
    "MATHG": "Mathematics",
    "HISTG": "History",
    "GEOGG": "Geography",
    "BIOLG": "Biology",
    "CHEMG": "Chemistry",
    "PHYSICSG": "Physics",
    "ECONG": "Economics",
    "PRINBG": "Principles of Business",
    "PRINAG": "Principles of Accounts",
    "SPANSG": "Spanish",
    "FRNCHG": "French",
    "ITG": "Information Technology",
    "ADDMTG": "Additional Mathematics",
    "OFFADG": "Office Administration",
    "AGSBG": "Agricultural Science (Double Award)",
    "AGSCG": "Agricultural Science (Single Award)",
    "SOCSG": "Social Studies",
    "INTSG": "Integrated Science",
    "HUMANG": "Human & Social Biology",
    "TDSCG": "Technical Drawing",
    "MECHTG": "Mechanical Engineering Technology",
    "FOODNG": "Food & Nutrition",
    "HTMG": "Hospitality Management",
    "ARTSG": "Visual Arts",
    "MUSCG": "Music",
    "DANCIG": "Dance",
    "THEATG": "Theatre Arts",
    "PHEDUG": "Physical Education",
    "CARITEG": "Caribbean History",
    # Additional mappings from original script
    "SOCSTUDG": "Social Studies",
    "OA": "Office Administration",
    "POAG": "Principles of Accounts",
    "HSBIOG": "Human and Social Biology",
    "POBG": "Principles of Business",
    "INTSCIG": "Integrated Science S/A",
    "BIOG": "Biology",
    "ADDMATH": "Additional Mathematics",
    "CARHISTG": "Caribbean History",
    "GEOG": "Geography",
    "SPANG": "Spanish",
    "FRENG": "French",
    "PORTG": "Portuguese",
    "EDPMG": "Electronic Document Preparation & Management",
    "RELIGEDG": "Religious Education",
    "TECHDRG": "Technical Drawing",
    "AGSCIDAG": "Agricultural Science D/A",
    "AGSCISAG": "Agricultural Science S/A",
    "INDTECHG": "Industrial Technology",
    "FASHION": "Textiles, Clothing & Fashion",
    "FOODNUTH": "Food, Nutrition & Health",
    "FAMRESMG": "Family & Resource Management",
    "MUSICG": "Music",
    "PEASPORT": "Physical Education & Sport",
    "THEARTSG": "Theatre Arts",
    "VISARTSG": "Visual Arts",
    "ACCU1": "Accounting Unit 1",
    "ACCU2": "Accounting Unit 2",
    "AMTU1": "Applied Mathematics Unit 1",
    "AMTU2": "Applied Mathematics Unit 2",
    "BIOU1": "Biology Unit 1",
    "BIOU2": "Biology Unit 2",
    "CARSTDU1": "Caribbean Studies",
    "CHEMU1": "Chemistry Unit 1",
    "CHEMU2": "Chemistry Unit 2",
    "COMMSTU1": "Communication Studies",
    "ECONU1": "Economics Unit 1",
    "ECONU2": "Economics Unit 2",
    "ENTRU1": "Entrepreneurship Unit 1",
    "ENTRU2": "Entrepreneurship Unit 2",
    "ENSCU1": "Environmental Science Unit 1",
    "ENSCU2": "Environmental Science Unit 2",
    "FRENU1": "French Unit 1",
    "FRENU2": "French Unit 2",
    "GEOU1": "Geography Unit 1",
    "GEOU2": "Geography Unit 2",
    "HISTU2": "History Unit 2",
    "INMATU1": "Integrated Mathematics",
    "INTHU1": "Information Technology Unit 1",
    "INTHU2": "Information Technology Unit 2",
    "LAWU1": "Law Unit 1",
    "LAWU2": "Law Unit 2",
    "LIEU1": "Literatures in English Unit 1",
    "LIEU2": "Literatures in English Unit 2",
    "MOBU1": "Management of Business Unit 1",
    "MOBU2": "Management of Business Unit 2",
    "PHYU1": "Physics Unit 1",
    "PHYU2": "Physics Unit 2",
    "PMATHU1": "Pure Mathematics Unit 1",
    "PMATHU2": "Pure Mathematics Unit 2",
    "SOCU1": "Sociology Unit 1",
    "SOCU2": "Sociology Unit 2",
    "SPU1": "Spanish Unit 1",
    "SPU2": "Spanish Unit 2",
    "TOURU1": "Tourism Unit 1",
    "TOURU2": "Tourism Unit 2"
}

SUBJECT_CODE_PATTERN = re.compile(r"([A-Z]{3,8})(?:-([A-Z]))?")  # Capture code and type separately
DATE_PATTERN = re.compile(r"\b(\d{2}/\d{2}/\d{4})\b")
CANDIDATE_NUM_PATTERN = re.compile(r"\b(\d{10})\b")
NAME_PATTERN = re.compile(r"^[A-Z'\- ]+,\s*[A-Z'\- ]+(?:\s+[A-Z'\- ]+)*$")


# ---------------- CSV PARSER (ROUTER) ----------------
def parse_csv(csv_path, exam_type, exam_month, log_callback):
    """Selects the appropriate CSV parser based on the exam type and month.

    This function acts as a router, directing the CSV parsing to the correct
    function based on whether the exam is a CSEC January exam or a standard
    May/June exam.

    Args:
        csv_path (str): The file path to the CSV file.
        exam_type (str): The type of exam (e.g., "CSEC", "CAPE").
        exam_month (str): The month of the exam (e.g., "January", "May - June").
        log_callback (function): A callback function for logging messages.

    Returns:
        list: A list of dictionaries, where each dictionary represents an
              eligible candidate.
    """
    if exam_type == "CSEC" and exam_month == "January":
        log_callback("Using CSEC January CSV parser.")
        return parse_csv_january(csv_path, log_callback)
    else:
        log_callback("Using standard May/June CSV parser.")
        return parse_csv_may_june(csv_path, exam_type, log_callback)


def parse_csv_may_june(csv_path, exam_type, log_callback):
    """Parses a CSV file for May/June CSEC and CAPE exams.

    This function reads a CSV file and extracts a list of candidates who are
    eligible for e-slips based on the services they have signed up for. It
    filters candidates by exam type to ensure only relevant candidates are
    processed.

    Args:
        csv_path (str): The file path to the CSV file.
        exam_type (str): The type of exam (e.g., "CSEC", "CAPE") to filter by.
        log_callback (function): A callback function for logging messages.

    Returns:
        list: A list of dictionaries, where each dictionary represents an
              eligible candidate. Returns an empty list if an error occurs.
    """
    eligible = []
    try:
        with open(csv_path, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f)
            log_callback(f"DEBUG: Reading {csv_path} for May/June...")
            log_callback(f"DEBUG: Headers found: {reader.fieldnames}")

            row_count = 0
            for row in reader:
                row_count += 1
                service = row.get("Additional Application Service - sent via email", "").strip()

                # --- DEBUG LOG ---
                log_callback(f"--- Processing Row {row_count} ---")
                log_callback(f"DEBUG: Found Service: '{service}'")

                if service not in [
                    "E-candidate slip/Timetable only- $30",
                    "Error recognition & E-candidate slip/Timetable- $50"
                ]:
                    # --- DEBUG LOG ---
                    log_callback("DEBUG: SKIPPING -> Service not eligible.")
                    continue

                if "Choose Examination" in row and row["Choose Examination"].strip():
                    exam_in_csv = row["Choose Examination"].strip().upper()
                    log_callback(f"DEBUG: Found Exam in CSV: '{exam_in_csv}'")
                    if exam_in_csv != exam_type.upper():
                        # --- DEBUG LOG ---
                        log_callback(f"DEBUG: SKIPPING -> Exam mismatch (App: {exam_type.upper()}, CSV: {exam_in_csv})")
                        continue

                last = (row.get("Last Name", "")).strip()
                first = (row.get("First Name", "")).strip()
                middle = (row.get("Middle Name", "")).strip()
                name = normalize_name_csv(last, first, middle)

                dob = normalize_dob(row.get("Date Of Birth", ""))
                if not dob:
                    log_callback(f"CSV row missing/invalid DOB for {name}; skipping")
                    continue

                # --- DEBUG LOG ---
                log_callback(f"DEBUG: SUCCESS -> Added candidate {name}")
                eligible.append({"name": name, "dob": dob, "raw": row})
        log_callback(f"CSV: {len(eligible)} candidate(s) eligible for e-slips after filtering.")
        return eligible
    except Exception as e:
        log_callback(f"ERROR parsing May/June CSV: {e}")
        return []


def parse_csv_january(csv_path, log_callback):
    """Parses a CSV file specifically for CSEC January exams.

    This parser is tailored for the format of CSEC January CSV files, which
    use different column headers for candidate names and services compared to
    the May/June exams. It extracts a list of eligible candidates based on the
    services they have signed up for.

    Args:
        csv_path (str): The file path to the CSV file.
        log_callback (function): A callback function for logging messages.

    Returns:
        list: A list of dictionaries, where each dictionary represents an
              eligible candidate. Returns an empty list if an error occurs.
    """
    eligible = []
    # Find the correct header names from the CSV
    name_col = "Full Name - name of candidate participating in CSEC January 2026 examination."
    service_col = "Application Processing Type - sent via email"
    dob_col = "Date of Birth"

    try:
        with open(csv_path, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f)

            # --- DEBUG LOG ---
            log_callback(f"DEBUG: Reading {csv_path} for January...")
            log_callback(f"DEBUG: Headers found: {reader.fieldnames}")

            # Check if headers exist
            if not reader.fieldnames:
                log_callback("ERROR: CSV file is empty or unreadable.")
                return []

            found_headers = {
                "name": any(name_col in h for h in reader.fieldnames),
                "service": any(service_col in h for h in reader.fieldnames),
                "dob": any(dob_col in h for h in reader.fieldnames)
            }

            # Dynamically find the full column names
            try:
                name_header = next(h for h in reader.fieldnames if name_col in h)
                log_callback(f"DEBUG: Found Name Header: '{name_header}'")  # --- DEBUG LOG ---
                service_header = next(h for h in reader.fieldnames if service_col in h)
                log_callback(f"DEBUG: Found Service Header: '{service_header}'")  # --- DEBUG LOG ---
                dob_header = next(h for h in reader.fieldnames if dob_col in h)
                log_callback(f"DEBUG: Found DOB Header: '{dob_header}'")  # --- DEBUG LOG ---
            except StopIteration:
                log_callback("ERROR: CSEC January CSV is missing required columns.")
                log_callback(f"  Missing Name Col ('{name_col}'): {not found_headers['name']}")  # More specific log
                log_callback(
                    f"  Missing Service Col ('{service_col}'): {not found_headers['service']}")  # More specific log
                log_callback(f"  Missing DOB Col ('{dob_col}'): {not found_headers['dob']}")  # More specific log
                return []

            row_count = 0  # --- DEBUG LOG ---
            for row in reader:
                row_count += 1  # --- DEBUG LOG ---
                service = row.get(service_header, "").strip()

                # --- DEBUG LOG ---
                log_callback(f"--- Processing Row {row_count} ---")
                log_callback(f"DEBUG: Found Service: '{service}'")

                # --- FIX 1: Added all valid service strings ---
                eligible_services = [
                    "Generate E-candidate slip/Timetable only- $30",
                    "Error recognition & E-candidate slip/Timetable- $50",
                    "E-candidate slip/Timetable only- $30",
                    "Error correction & E-candidate slip/Timetable- $50"  # Added this from your log
                ]

                if service not in eligible_services:
                    log_callback("DEBUG: SKIPPING -> Service not eligible.")  # --- DEBUG LOG ---
                    continue

                full_name_str = row.get(name_header, "").strip()
                name = normalize_name_from_full(full_name_str)

                dob = normalize_dob(row.get(dob_header, ""))
                if not dob:
                    log_callback(f"CSV row missing/invalid DOB for {name}; skipping")
                    continue

                log_callback(f"DEBUG: SUCCESS -> Added candidate {name}")  # --- DEBUG LOG ---
                eligible.append({"name": name, "dob": dob, "raw": row})
        log_callback(f"CSV: {len(eligible)} candidate(s) eligible for e-slips after filtering.")
        return eligible
    except Exception as e:
        log_callback(f"ERROR parsing CSEC January CSV: {e}")
        return []


# ---------------- PDF TEXT HELPERS ----------------
def extract_text_from_pdf(pdf_path, log, output_dir):
    """Extracts text from a PDF file, prioritizing OCR for accuracy.

    This function processes each page of a PDF file, using image-based OCR
    (Optical Character Recognition) to extract text. This method is generally
    more accurate for scanned documents than simple text extraction. The
    extracted text from all pages is concatenated and returned as a single
    string.

    Args:
        pdf_path (str): The file path to the PDF file.
        log (function): A callback function for logging messages.
        output_dir (str): The directory to save the OCR output text file.

    Returns:
        str: The combined, cleaned text extracted from the PDF.
    """
    doc = pymupdf.open(pdf_path)
    all_text = []
    page_count = len(doc)
    pages_to_process = page_count if not DEBUG else 1

    full_ocr_text = ""

    for i, page in enumerate(doc):
        if DEBUG and i >= pages_to_process:
            break

        log(f"Processing page {i + 1}/{pages_to_process}")

        txt = ""
        if pytesseract and Image:
            log(f"Forcing high-DPI OCR on page {i + 1}...")
            pix = page.get_pixmap(dpi=350)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            config = '--psm 6'
            txt = pytesseract.image_to_string(img, config=config)
        else:
            log("Tesseract/Pillow not found, falling back to simple text extraction.")
            txt = page.get_text("text") or ""

        cleaned_txt = clean_ocr_text(txt)
        all_text.append(cleaned_txt)
        full_ocr_text += f"--- PAGE {i + 1} ---\n{cleaned_txt}\n\n"

    doc.close()

    if output_dir and full_ocr_text.strip():
        output_filepath = os.path.join(output_dir, f"ocr_output_{os.path.basename(pdf_path)}.txt")
        try:
            with open(output_filepath, "w", encoding="utf-8") as f:
                f.write(full_ocr_text)
            log(f"Saved OCR text to {os.path.basename(output_filepath)}")
        except Exception as e:
            log(f"Could not save OCR text: {e}")

    joined = "\n".join(all_text)
    log(f"Extracted text from {os.path.basename(pdf_path)} (pages processed: {pages_to_process}).")
    return joined


def clean_ocr_text(s):
    """Cleans and corrects common errors in OCR-extracted text.

    This function performs a series of string replacements to fix common
    misinterpretations that occur during Optical Character Recognition (OCR),
    such as replacing visually similar characters (e.g., 'O' with '0') and
    removing unwanted symbols.

    Args:
        s (str): The input string from OCR extraction.

    Returns:
        str: The cleaned and corrected string.
    """
    s = s.replace("\u2010", "-").replace("\u2011", "-").replace("\u2013", "-")
    s = s.replace('|', '').replace(']', '')

    s = s.replace(' O ', ' 0 ').replace(' G ', ' 6 ').replace(' B ', ' 8 ')
    s = s.replace(' l ', ' 1 ')

    s = re.sub(r'\s+', ' ', s)
    return s


def normalize_name_csv(last, first, middle):
    """Normalizes a candidate's name from separate parts into a standard format.

    This function takes the last, first, and middle names as separate strings,
    trims whitespace, converts them to a consistent case, and combines them
    into a single string in the format "Last, First Middle".

    Args:
        last (str): The candidate's last name.
        first (str): The candidate's first name.
        middle (str): The candidate's middle name.

    Returns:
        str: The normalized full name string.
    """
    parts = []
    last = last.strip().upper()
    first = first.strip().upper()
    middle = middle.strip().upper()
    if last:
        nm = last + ", " + first
        if middle:
            nm += " " + middle
        parts.append(nm)
    else:
        nm = first
        if middle:
            nm += " " + middle
        parts.append(nm)
    return ", ".join(parts).title()


def normalize_name_from_full(full_name_str):
    """Parses a full name string into the "Last, First Middle" format.

    This function is designed to handle full names provided as a single string
    (e.g., "First Middle Last") and reformat them into a standardized
    "Last, First Middle" format.

    Args:
        full_name_str (str): The full name of the candidate.

    Returns:
        str: The normalized name string, or an empty string if the input is
             empty.
    """
    parts = full_name_str.strip().upper().split()
    if not parts:
        return ""

    last = parts[-1]
    first = parts[0] if len(parts) > 1 else ""
    middle = " ".join(parts[1:-1])

    return normalize_name_csv(last, first, middle)  # Re-use existing formatter


def normalize_key_name(name_str):
    """Creates a simplified, consistent key from a name string for matching.

    This function converts a name string into a format that is reliable for
    comparisons and lookups. It removes all whitespace and commas and converts
    the string to lowercase to ensure that variations in formatting do not
    affect matching.

    Args:
        name_str (str): The name string to normalize.

    Returns:
        str: The normalized key string, or an empty string if the input is not
             a string.
    """
    if not isinstance(name_str, str):
        return ""
    return re.sub(r'[\s,]+', '', name_str).lower()


def normalize_dob(val):
    """Normalizes a date of birth string into a 'dd/mm/yyyy' format.

    This function attempts to parse a date of birth from a string, which may
    be in one of several common formats (e.g., 'yyyy-mm-dd', 'dd/mm/yyyy',
    'Month dd, yyyy'). It returns the date formatted as 'dd/mm/yyyy'.

    Args:
        val (str): The date of birth string.

    Returns:
        str or None: The normalized date string, or None if the input cannot
                     be parsed.
    """
    val = str(val).strip()
    if not val:
        return None

    val = val.replace(".", "").strip()

    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue

    extended_formats = [
        "%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y",
        "%b %d %Y", "%B %d %Y", "%d-%b-%Y", "%d-%B-%Y",
        "%b-%d-%Y", "%B-%d-%Y", "%d %b, %Y", "%d %B, %Y",
    ]

    for fmt in extended_formats:
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    try:
        words = val.split()
        if len(words) >= 3:
            words[0] = words[0].capitalize()
            val_capitalized = " ".join(words)
            for fmt in extended_formats:
                try:
                    dt = datetime.strptime(val_capitalized, fmt)
                    return dt.strftime("%d/%m/%Y")
                except ValueError:
                    continue
    except Exception:
        pass

    m = DATE_PATTERN.search(val)
    return m.group(1) if m else None


# ---------------- CENTRE LIST PARSER (REVISED) ----------------
def parse_centre_list(pdf_path, log, output_dir):
    """Parses a PDF file to extract a list of exam centres.

    This function reads a PDF centre list, extracts the text, and uses regular
    expressions to identify 6-digit centre codes and their corresponding names.
    It includes logic to clean up common OCR errors and correctly associate
    names with codes.

    Args:
        pdf_path (str): The file path to the centre list PDF.
        log (function): A callback function for logging messages.
        output_dir (str): The directory to save the OCR output text file.

    Returns:
        dict: A dictionary mapping centre codes (str) to centre names (str).
              Returns an empty dictionary if parsing fails.
    """
    centres = {}
    log("--- STARTING CENTRE LIST PARSING (IMPROVED LOGIC) ---")
    try:
        text = extract_text_from_pdf(pdf_path, log, output_dir)

        # This new logic is more robust for messy OCR that merges lines.
        # 1. Find all 6-digit codes in the text to use as anchors.
        # 2. The name of a centre is the text between its code and the next code.
        # 3. Clean the extracted text to remove area names and other junk.

        code_pattern = re.compile(r"\b(\d{6})\b")
        matches = list(code_pattern.finditer(text))

        if not matches:
            log("Warning: No 6-digit centre codes found in the centre list PDF.")
            return {}

        for i, current_match in enumerate(matches):
            code = current_match.group(1)

            # Define the slice of text that contains the name
            start_pos = current_match.end()
            end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(text)

            name_raw = text[start_pos:end_pos]

            # Clean the raw name string
            name_cleaned = re.sub(r'\d', '', name_raw)  # Remove stray digits
            name_cleaned = name_cleaned.replace("E-Testing", "").strip()
            name_cleaned = name_cleaned.replace("Schoo!", "School").strip()  # Fix common OCR error

            # Intelligently find the end of the school name before the area name starts
            common_endings = ["Secondary School", "Secondary", "College", "Campus", "High School", "School"]

            best_name = name_cleaned
            # We iterate through possible name endings and take the longest, most complete name we can find.
            for ending in common_endings:
                if ending in name_cleaned:
                    end_index = name_cleaned.find(ending) + len(ending)
                    potential_name = name_cleaned[:end_index].strip()
                    best_name = potential_name
                    break  # Found a good candidate, stop here.

            # Final cleanup
            best_name = re.sub(r'\s+', ' ', best_name).strip()

            if code and best_name:
                log(f"  Parsed Centre: {code} -> '{best_name}'")
                centres[code] = best_name

        if not centres:
            log("Warning: No centres parsed from centre list PDF with improved logic.")
        else:
            log(f"Centres parsed (improved logic): {len(centres)}")
        return centres

    except Exception as e:
        log(f"ERROR parsing centre list: {e}")
        return {}


# ---------------- CANDIDATE LIST PARSER (REVISED) ----------------
def parse_candidate_list(pdf_path, log, output_dir):
    """Parses a PDF candidate list to extract detailed candidate information.

    This function reads a PDF file containing a list of candidates and their
    details, including candidate number, name, date of birth, gender, and
    enrolled subjects. It uses pattern matching to identify and extract this
    information for each candidate.

    Args:
        pdf_path (str): The file path to the candidate list PDF.
        log (function): A callback function for logging messages.
        output_dir (str): The directory to save the OCR output text file.

    Returns:
        tuple: A tuple containing:
            - list: A list of dictionaries, where each dictionary represents a
                    candidate.
            - list: A list of any text blocks that could not be parsed.
            - str: The full, cleaned text extracted from the PDF.
    """
    candidates = []
    log("--- STARTING CANDIDATE LIST PARSING (Smarter Logic V3) ---")

    try:
        text = extract_text_from_pdf(pdf_path, log, output_dir)

        matches = list(re.finditer(CANDIDATE_NUM_PATTERN, text))

        if not matches:
            raise ValueError("OCR did not find any 10-digit candidate numbers in the PDF text.")

        for i, current_match in enumerate(matches):
            start_pos = current_match.start()
            end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(text)

            block = text[start_pos:end_pos]
            cleaned_block = re.sub(r'\s+', ' ', block).strip()

            id_match = CANDIDATE_NUM_PATTERN.search(cleaned_block)
            dob_match = DATE_PATTERN.search(cleaned_block)

            if not (id_match and dob_match):
                log(f"SKIPPING malformed block: {cleaned_block[:100]}...")
                continue

            candidate_num_full = id_match.group(1)
            dob = dob_match.group(1)

            name_raw = cleaned_block[id_match.end():dob_match.start()].strip()
            name_parts = [part.strip() for part in name_raw.split(',') if part.strip()]
            if name_parts:
                last_name = name_parts[0]
                first_middle = ' '.join(name_parts[1:])
                name = f"{last_name}, {first_middle}".title()
            else:
                name = name_raw.title()

            if not name or not re.search(r'[a-zA-Z]', name):
                log(f"SKIPPING block for Cand# {candidate_num_full}: Invalid name parsed ('{name_raw}').")
                continue

            remaining_text = cleaned_block[dob_match.end():].strip()

            gender_match = re.search(r'\b([MF])\b', remaining_text)
            if not gender_match:
                log(f"SKIPPING block for Cand# {candidate_num_full}: Could not find Gender after DOB.")
                continue

            gender = "Male" if gender_match.group(1) == "M" else "Female"

            subjects_raw = remaining_text[gender_match.end():].strip()

            count_match = re.search(r'\s(\d)$', subjects_raw)
            if count_match:
                subjects_raw = subjects_raw[:count_match.start()].strip()

            subjects_list = []
            for code_match in SUBJECT_CODE_PATTERN.finditer(subjects_raw.upper()):
                code, type = code_match.groups()
                # MODIFICATION: Only add subject if the code is a valid key in our map
                if code in SUBJECT_CODE_MAP:
                    subjects_list.append({"code": code, "type": type or 'N/A'})

            candidates.append({
                "id": candidate_num_full, "centre_num": candidate_num_full[:6], "seq_num": candidate_num_full[6:],
                "name": name, "dob": dob, "gender": gender, "subjects": subjects_list
            })

        log(f"Found {len(candidates)} candidates successfully parsed.")

        if not candidates:
            raise ValueError("OCR parsing failed to extract any valid candidate data.")

        return candidates, [], text

    except Exception as e:
        log(f"ERROR parsing candidate list: {e}")
        return [], [], ""


# ---------------- MANUAL WINDOWS ----------------
# Helper class for scrollable frames
class ScrollableFrame(ttk.Frame):
    """A custom tkinter frame that includes a vertical scrollbar.

    This class creates a scrollable container, which is useful for displaying
    content that may exceed the visible area of a window. It combines a
    `ttk.Frame` with a `tk.Canvas` and a `ttk.Scrollbar` to achieve the
    scrolling effect.
    """

    def __init__(self, container, *args, **kwargs):
        """Initializes the ScrollableFrame.

        Args:
            container (tk.Widget): The parent widget for this frame.
            *args: Variable length argument list.
            **kwargs: Arbitrary keyword arguments.
        """
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


class BaseManualEntry(tk.Toplevel):
    """A base class for creating manual data entry dialog windows.

    This class provides the basic structure for a top-level dialog window,
    including a title, instructions, and Submit/Cancel buttons. It is intended
    to be subclassed for specific data entry tasks.
    """

    def __init__(self, parent, title, instructions):
        """Initializes the BaseManualEntry dialog.

        Args:
            parent (tk.Widget): The parent widget for this dialog.
            title (str): The title of the dialog window.
            instructions (str): The instructions to display at the top of the
                                dialog.
        """
        super().__init__(parent)
        self.title(title)
        self.geometry("950x600")  # Increased size for tables and new button
        self.transient(parent)
        self.grab_set()
        self.entries = []

        self.main_frame = ttk.Frame(self, padding=10)
        self.main_frame.pack(fill="both", expand=True)

        self.instructions_label = ttk.Label(self.main_frame, text=instructions, wraplength=930)
        self.instructions_label.pack(pady=(0, 8), fill='x')

        btns = ttk.Frame(self)
        btns.pack(pady=8)
        ttk.Button(btns, text="Submit", command=self.submit).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left")

    def submit(self):
        """Handles the submission of the entered data.

        This method is intended to be overridden by subclasses to provide
        specific logic for processing and saving the data entered in the
        dialog.
        """
        raise NotImplementedError

    def show(self):
        """Displays the dialog window and waits for it to be closed.

        This method makes the dialog window visible and waits until the user
        interacts with it (e.g., by clicking Submit or Cancel).

        Returns:
            list: The entries collected from the dialog.
        """
        self.deiconify()
        self.wait_window()
        return self.entries


class ManualCandidateEntry(BaseManualEntry):
    """A dialog for manually entering or correcting candidate information.

    This class provides a user interface for manually adding or editing
    candidate data that could not be automatically parsed from the PDF. It
    includes a table for entering candidate details and a button to auto-fill
    information by searching the PDF text.
    """

    def __init__(self, parent, missed_blocks, pdf_text):
        """Initializes the ManualCandidateEntry dialog.

        Args:
            parent (tk.Widget): The parent widget for this dialog.
            missed_blocks (list): A list of text blocks that could not be
                                  parsed.
            pdf_text (str): The full text extracted from the candidate list
                            PDF.
        """
        super().__init__(parent, "Manual Candidate Entry",
                         "Review any automatically extracted but unparsable lines below. Add or correct candidate information in the table. Use the 'Find All Details' button to search the PDF and auto-fill details for all rows with a valid Candidate ID.")

        self.pdf_text = pdf_text  # NEW: Store full PDF text for searching
        self.rows = []

        if missed_blocks:
            ref_frame = ttk.LabelFrame(self.main_frame, text="Unparsable Lines (for reference)", padding=5)
            ref_frame.pack(fill='x', padx=10, pady=(0, 5))
            ref_text = tk.Text(ref_frame, height=4, font=("Courier New", 9))
            ref_text.pack(fill='x', expand=True)
            ref_text.insert("1.0", "\n---\n".join(missed_blocks))
            ref_text.config(state="disabled")

        self.scrollable_frame = ScrollableFrame(self.main_frame)
        self.scrollable_frame.pack(fill="both", expand=True, padx=10)
        self.container = self.scrollable_frame.scrollable_frame

        self.headers = ["Candidate ID", "Full Name (Surname, First)", "DOB (dd/mm/yyyy)", "Gender (M/F)",
                        "Subjects (e.g., MATHG-A POBG-R)"]  # MODIFIED: Removed Actions column
        for i, header in enumerate(self.headers):
            ttk.Label(self.container, text=header, font=("Helvetica", 10, "bold")).grid(row=0, column=i, padx=5, pady=5,
                                                                                        sticky='w')

        # NEW: Bottom button frame
        bottom_btn_frame = ttk.Frame(self)
        bottom_btn_frame.pack(pady=5, padx=10, fill='x')

        find_all_btn = ttk.Button(bottom_btn_frame, text="Find All Details", command=self._find_all_details)
        find_all_btn.pack(side=tk.LEFT)

        add_btn = ttk.Button(bottom_btn_frame, text="Add New Candidate Row", command=self._add_row)
        add_btn.pack(side=tk.RIGHT)

        self._add_row()  # Start with one empty row

    def _add_row(self, data=None):
        """Adds a new row to the candidate entry table.

        This method creates a new set of entry widgets for a single candidate
        and adds them to the table. It can optionally populate the new row with
        initial data.

        Args:
            data (dict, optional): A dictionary containing initial data for the
                                 new row. Defaults to None.
        """
        data = data or {}
        row_num = len(self.rows) + 1

        row_vars = {
            "id": tk.StringVar(value=data.get("id", "")),
            "name": tk.StringVar(value=data.get("name", "")),
            "dob": tk.StringVar(value=data.get("dob", "")),
            "gender": tk.StringVar(value=data.get("gender", "")),
            "subjects": tk.StringVar(value=data.get("subjects", ""))
        }

        ttk.Entry(self.container, textvariable=row_vars["id"], width=15).grid(row=row_num, column=0, padx=5, pady=2,
                                                                              sticky='w')
        ttk.Entry(self.container, textvariable=row_vars["name"], width=30).grid(row=row_num, column=1, padx=5, pady=2,
                                                                                sticky='w')
        ttk.Entry(self.container, textvariable=row_vars["dob"], width=15).grid(row=row_num, column=2, padx=5, pady=2,
                                                                               sticky='w')
        ttk.Entry(self.container, textvariable=row_vars["gender"], width=5).grid(row=row_num, column=3, padx=5, pady=2,
                                                                                 sticky='w')
        ttk.Entry(self.container, textvariable=row_vars["subjects"], width=40).grid(row=row_num, column=4, padx=5,
                                                                                    pady=2, sticky='w')

        self.rows.append(row_vars)

    # NEW: Method to find details for all candidates in the table
    def _find_all_details(self):
        """Finds and auto-fills details for all candidates in the table.

        This method iterates through all rows in the candidate entry table. For
        each row that has a valid 10-digit candidate ID, it searches the full
        PDF text for that ID and, if found, extracts and populates the
        gender and subjects fields.
        """
        found_count = 0
        not_found_ids = []

        for row_vars in self.rows:
            candidate_id = row_vars["id"].get().strip()
            if not re.match(r'^\d{10}$', candidate_id):
                continue  # Skip rows without a valid ID

            matches = list(re.finditer(CANDIDATE_NUM_PATTERN, self.pdf_text))
            found_block = None

            for i, current_match in enumerate(matches):
                if current_match.group(1) == candidate_id:
                    start_pos = current_match.start()
                    end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(self.pdf_text)
                    found_block = self.pdf_text[start_pos:end_pos]
                    break

            if not found_block:
                not_found_ids.append(candidate_id)
                continue

            cleaned_block = re.sub(r'\s+', ' ', found_block).strip()
            gender_match = re.search(r'\b([MF])\b', cleaned_block)

            if not gender_match:
                not_found_ids.append(candidate_id)
                continue

            gender = gender_match.group(1)
            subjects_raw = cleaned_block[gender_match.end():].strip()
            count_match = re.search(r'\s(\d)$', subjects_raw)
            if count_match:
                subjects_raw = subjects_raw[:count_match.start()].strip()

            subjects_list = []
            for code_match in SUBJECT_CODE_PATTERN.finditer(subjects_raw.upper()):
                code, type = code_match.groups()
                if code in SUBJECT_CODE_MAP:
                    subject_str = code
                    if type:
                        subject_str += f"-{type}"
                    subjects_list.append(subject_str)

            subjects_final_str = " ".join(subjects_list)

            if not subjects_final_str:
                not_found_ids.append(candidate_id)
                continue

            row_vars["gender"].set(gender)
            row_vars["subjects"].set(subjects_final_str)
            found_count += 1

        summary_message = f"Found and populated details for {found_count} candidate(s).\n\n"
        if not_found_ids:
            summary_message += f"Could not find or fully parse details for the following IDs:\n" + "\n".join(
                not_found_ids)

        messagebox.showinfo("Find Complete", summary_message, parent=self)

    def submit(self):
        """Processes and validates the manually entered candidate data.

        This method is called when the user clicks the "Submit" button. It
        iterates through each row in the table, validates the entered data
        (e.g., candidate ID format, DOB format), and compiles a list of valid
        candidate dictionaries.
        """
        out = []
        has_blank_id_with_data = False

        for row_vars in self.rows:
            id_val = row_vars["id"].get().strip()
            name_val = row_vars["name"].get().strip()
            dob_val = row_vars["dob"].get().strip()
            if not id_val and (name_val or dob_val):
                has_blank_id_with_data = True
                break

        if has_blank_id_with_data:
            proceed = messagebox.askyesno(
                "Blank Candidate ID",
                "One or more rows have a blank Candidate ID but contain other data. These incomplete rows will be ignored.\n\nDo you wish to continue?"
            )
            if not proceed:
                return

        for row_vars in self.rows:
            id_val = row_vars["id"].get().strip()
            name_val = row_vars["name"].get().strip()
            dob_val = row_vars["dob"].get().strip()
            gender_val = row_vars["gender"].get().strip().upper()
            subjects_str = row_vars["subjects"].get().strip().upper()

            if not id_val:
                continue

            if not re.match(r'^\d{10}$', id_val):
                messagebox.showerror("Validation Error", f"Invalid Candidate ID: '{id_val}'. Must be 10 digits.")
                return

            normalized_dob = normalize_dob(dob_val)
            if not normalized_dob:
                messagebox.showerror("Validation Error", f"Invalid DOB: '{dob_val}'. Use dd/mm/yyyy format.")
                return

            if gender_val and gender_val not in ["M", "F"]:
                messagebox.showerror("Validation Error", f"Invalid Gender: '{gender_val}'. Use M or F.")
                return

            gender_full = "Male" if gender_val == "M" else "Female" if gender_val == "F" else "N/A"

            subjects_list = []
            for s in subjects_str.split():
                match = SUBJECT_CODE_PATTERN.match(s)
                if match:
                    code, type = match.groups()
                    subjects_list.append({"code": code, "type": type or 'N/A'})

            out.append({
                "id": id_val, "centre_num": id_val[:6], "seq_num": id_val[6:],
                "name": name_val, "dob": normalized_dob, "gender": gender_full,
                "subjects": subjects_list
            })

        self.entries = out
        self.destroy()


class ManualCSVEntry(ManualCandidateEntry):
    """A dialog for matching unmatched CSV candidates with PDF data.

    This class provides a user interface for resolving candidates who were
    present in the eligibility CSV but not found in the candidate list PDF.
    It pre-populates the entry table with the names and DOBs from the CSV,
    allowing the user to manually enter the correct 10-digit candidate ID.
    """

    def __init__(self, parent, unmatched_csv, pdf_text):
        """Initializes the ManualCSVEntry dialog.

        Args:
            parent (tk.Widget): The parent widget for this dialog.
            unmatched_csv (list): A list of candidate dictionaries from the CSV
                                  that were not found in the PDF.
            pdf_text (str): The full text extracted from the candidate list PDF.
        """
        # MODIFIED: Pass pdf_text to the parent class
        super().__init__(parent, [], pdf_text)
        self.title("Manual CSV Candidate Entry")
        self.instructions_label.config(
            text="The following candidates from your CSV were not found in the PDF. Please enter their 10-digit Candidate ID, use the 'Find All Details' button to auto-populate subjects, then click Submit.")

        for widget in self.container.winfo_children():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        self.rows = []

        for i, header in enumerate(self.headers):
            ttk.Label(self.container, text=header, font=("Helvetica", 10, "bold")).grid(row=0, column=i, padx=5, pady=5,
                                                                                        sticky='w')

        for c in unmatched_csv:
            data = {"id": "??????????", "name": c['name'], "dob": c['dob']}
            self._add_row(data)


class ManualCentreEntry(BaseManualEntry):
    """A dialog for manually entering the names of missing exam centres.

    This class provides a user interface for the user to input the names for
    any centre codes that were found in the candidate list but not in the
    centre list PDF.
    """

    def __init__(self, parent, missing_centre_codes):
        """Initializes the ManualCentreEntry dialog.

        Args:
            parent (tk.Widget): The parent widget for this dialog.
            missing_centre_codes (list): A list of centre codes for which the
                                         names are missing.
        """
        super().__init__(parent, "Manual Centre Entry",
                         "Enter the name for each missing centre code. Rows with empty codes or names will be ignored.")
        self.centre_entries = {}

        scrollable_frame = ScrollableFrame(self.main_frame)
        scrollable_frame.pack(fill="both", expand=True, padx=10)
        container = scrollable_frame.scrollable_frame

        ttk.Label(container, text="Centre Code", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=5, pady=5,
                                                                                      sticky='w')
        ttk.Label(container, text="Centre Name", font=("Helvetica", 10, "bold")).grid(row=0, column=1, padx=5, pady=5,
                                                                                      sticky='w')

        for i, code in enumerate(missing_centre_codes, 1):
            ttk.Label(container, text=code).grid(row=i, column=0, padx=5, pady=2, sticky='w')
            name_var = tk.StringVar()
            name_entry = ttk.Entry(container, textvariable=name_var, width=50)
            name_entry.grid(row=i, column=1, padx=5, pady=2, sticky='w')
            self.centre_entries[code] = name_var

    def submit(self):
        """Processes and saves the manually entered centre names.

        This method is called when the user clicks the "Submit" button. It
        collects the entered names for each centre code and stores them in a
        dictionary.
        """
        centres = {}
        for code, name_var in self.centre_entries.items():
            name = name_var.get().strip()
            if code and name:
                centres[code] = name
        self.entries = centres
        self.destroy()


class ManualTimetableEntry(BaseManualEntry):
    """A dialog for manually entering exam timetable information.

    This class provides a user interface for entering the date, time, and
    paper number for each subject in the exam. This information is used to
    populate the timetable section of the e-slips.
    """

    def __init__(self, parent, unique_subjects, exam_month, exam_year):
        """Initializes the ManualTimetableEntry dialog.

        Args:
            parent (tk.Widget): The parent widget for this dialog.
            unique_subjects (list): A list of unique subject codes that require
                                    timetable information.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
        """
        super().__init__(parent, "Manual Timetable Entry",
                         "Enter timetable information for each paper of a subject. Use the 'Add Paper' button if a subject has more than three papers.")
        self.subject_rows = {}

        scrollable_frame = ScrollableFrame(self.main_frame)
        scrollable_frame.pack(fill="both", expand=True, padx=10)
        container = scrollable_frame.scrollable_frame

        self.period_options = ["AM", "PM", "Oral Examination"]

        for s_code in sorted(unique_subjects):
            s_name = SUBJECT_CODE_MAP.get(s_code, s_code)

            subject_frame = ttk.LabelFrame(container, text=f"{s_code} - {s_name}", padding=10)
            subject_frame.pack(fill="x", pady=5, padx=5)

            headers = ["Paper", "Date", "Period"]
            for i, header in enumerate(headers):
                ttk.Label(subject_frame, text=header, font=("Helvetica", 10, "bold")).grid(row=0, column=i, padx=5,
                                                                                           pady=5)

            self.subject_rows[s_code] = []

            default_papers = ["1", "2", "3/2"]
            for paper_placeholder in default_papers:
                self._add_paper_row(s_code, subject_frame, paper_placeholder)

            add_button = ttk.Button(subject_frame, text=f"Add Paper for {s_code}",
                                    command=lambda code=s_code, frame=subject_frame: self._add_paper_row(code, frame))
            add_button.grid(row=len(self.subject_rows[s_code]) + 1, column=0, columnspan=3, pady=5)

    def _add_paper_row(self, subject_code, container, paper_num=None):
        """Adds a new row for a single paper to the timetable entry table.

        This method creates the entry widgets for a single paper (paper number,
        date, and session) and adds them to the appropriate subject's frame.

        Args:
            subject_code (str): The subject code for which to add a paper.
            container (tk.Widget): The parent widget for the new row.
            paper_num (str, optional): The default paper number to display.
                                     Defaults to None.
        """
        row_index = len(self.subject_rows[subject_code]) + 1

        paper_var = tk.StringVar(value=str(paper_num) if paper_num else str(row_index))
        date_var = tk.StringVar()
        period_var = tk.StringVar(value="AM")

        ttk.Entry(container, textvariable=paper_var, width=10).grid(row=row_index, column=0, padx=5, pady=2)
        ttk.Entry(container, textvariable=date_var, width=20).grid(row=row_index, column=1, padx=5, pady=2)
        ttk.Combobox(container, textvariable=period_var, values=self.period_options, width=20, state="readonly").grid(
            row=row_index, column=2, padx=5, pady=2)

        self.subject_rows[subject_code].append((paper_var, date_var, period_var))

    def submit(self):
        """Processes and saves the manually entered timetable information.

        This method is called when the user clicks the "Submit" button. It
        collects the entered timetable data for each subject and stores it in a
        dictionary.
        """
        timetable = {}
        for subject_code, rows in self.subject_rows.items():
            timetable[subject_code] = []
            for paper_var, date_var, period_var in rows:
                date = date_var.get().strip()
                if date:  # Only add rows that have a date
                    timetable[subject_code].append({
                        "paper": paper_var.get().strip(),
                        "date": date,
                        "session": period_var.get().strip()
                    })
        self.entries = timetable
        self.destroy()


# --- PDF GENERATION CLASS ---
class PDF(FPDF):
    """A customized PDF class for generating e-slips with a background image.

    This class extends the `FPDF` library to include a background image on each
    page and to define a custom header and footer for the e-slips.
    """

    def __init__(self, background_path="background.png"):
        """Initializes the PDF object.

        Args:
            background_path (str, optional): The file path to the background
                                             image. Defaults to "background.png".
        """
        super().__init__()
        self.background_path = background_path

    def add_page(self, orientation='', format='', same=False):
        """Adds a new page to the PDF and applies the background image.

        Overrides the default `add_page` method to include the functionality
        of adding a background image to each new page.
        """
        super().add_page(orientation, format, same)
        if self.background_path and os.path.exists(self.background_path):
            self.image(self.background_path, x=0, y=0, w=self.w, h=self.h)
        else:
            try:
                self.set_font('Roboto', 'I', 8)
                self.cell(0, 5, "background.png not found", 0, align='C')
            except RuntimeError:
                self.set_font('Helvetica', 'I', 8)
                self.cell(0, 5, "background.png not found", 0, align='C')

    def header(self):
        """Defines the header section of the PDF page.

        This method is automatically called by `FPDF` when a new page is
        created. It sets the font and adds the main title to the e-slip.
        """
        try:
            self.set_font('Roboto', 'B', 14)
        except RuntimeError:
            self.set_font('Helvetica', 'B', 14)
        self.set_y(50)
        self.cell(0, 10, 'CXC Examination E-Slip', 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def footer(self):
        """Defines the footer section of the PDF page.

        This method is automatically called by `FPDF` and is used to position
        the page content from the bottom of the page.
        """
        self.set_y(-15)


def create_pdf_slip(candidate, centre_name, timetable, output_dir, exam_month, exam_year, exam_type):
    """Generates a single PDF e-slip for a given candidate.

    This function takes all the necessary information for a single candidate and
    creates a formatted PDF e-slip. The e-slip includes the candidate's
    credentials, exam details, and a personalized timetable.

    Args:
        candidate (dict): A dictionary containing the candidate's information.
        centre_name (str): The name of the exam centre.
        timetable (dict): A dictionary containing the timetable information for
                        all subjects.
        output_dir (str): The directory where the generated PDF will be saved.
        exam_month (str): The month of the exam.
        exam_year (str): The year of the exam.
        exam_type (str): The type of exam (e.g., "CSEC", "CAPE").

    Returns:
        str or bool: The file path of the generated PDF if successful,
                     otherwise False.
    """
    try:
        name_parts = candidate['name'].split(',', 1)
        surname = name_parts[0].strip() if name_parts else "Unknown"
        other_names = name_parts[1].strip() if len(name_parts) > 1 else ""

        sanitized_surname = re.sub(r'[\\/*?:"<>|]', "", surname)
        sanitized_other_names = re.sub(r'[\\/*?:"<>|]', "", other_names)

        filename = f"{exam_type} E-Slip {sanitized_surname} {sanitized_other_names}.pdf"
        filepath = os.path.join(output_dir, filename)

        counter = 1
        base_filepath = filepath
        while os.path.exists(filepath):
            name, ext = os.path.splitext(base_filepath)
            filepath = f"{name} ({counter}){ext}"
            counter += 1

        pdf = PDF()

        try:
            pdf.add_font('Roboto', '', 'Roboto-Regular.ttf')
            pdf.add_font('Roboto', 'B', 'Roboto-Bold.ttf')
            pdf.add_font('Roboto', 'I', 'Roboto-Italic.ttf')
        except Exception as e:
            if "FPDFException" in str(e) and "TTF file" in str(e):
                messagebox.showerror("Font File Missing",
                                     "Required Roboto font files (.ttf) not found. Please place Roboto-Regular.ttf, Roboto-Bold.ttf, and Roboto-Italic.ttf in the same directory as the script.")
                return False

        pdf.add_page()
        pdf.set_right_margin(15)
        pdf.set_left_margin(15)
        pdf.set_auto_page_break(True, 30)

        pdf.set_font('Roboto', '', 11)
        greeting_text = (
            "Greetings\n"
            f"Thank you for using the Exam Concierge Service for the {exam_month} {exam_year} CXC examinations.\n"
            "The information requested is as follows:"
        )
        pdf.multi_cell(0, 5, greeting_text)
        pdf.ln(5)

        pdf.set_font('Roboto', 'B', 12)
        pdf.cell(0, 10, 'Candidate Credentials', 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        pdf.set_font('Roboto', '', 10)
        label_col_width = 55
        value_col_width = 125

        credentials = {
            "Candidate Surname": surname, "Candidate First/Other Names": other_names,
            "Candidate DOB": candidate['dob'], "Candidate Gender": candidate['gender'],
            "Candidate Number": candidate['id'], "Centre Number": candidate['centre_num'],
            "Centre Location": centre_name or 'N/A'  # Handle empty centre name gracefully
        }
        pdf.set_fill_color(220, 220, 220)
        for label, value in credentials.items():
            pdf.set_font('Roboto', 'B', 10)
            pdf.cell(label_col_width, 8, label, border=1, fill=True)
            pdf.set_font('Roboto', '', 10)
            pdf.cell(value_col_width, 8, f" {value}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(5)

        examination_details = f"{exam_type} {exam_month} {exam_year}"
        pdf.set_font('Roboto', 'B', 10)
        pdf.cell(label_col_width, 8, "Examination", border=1, fill=True)
        pdf.set_font('Roboto', '', 10)
        pdf.cell(0, 8, f" {examination_details}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(2)

        pdf.set_font('Roboto', 'B', 10)
        pdf.set_fill_color(220, 220, 220)
        pdf.cell(0, 8, "Examination Timetable", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)

        pdf.set_font('Roboto', 'B', 10)
        headers = ["Subject", "Candidate Type", "Paper", "Date", "Session"]
        total_width = pdf.w - pdf.l_margin - pdf.r_margin
        col_widths = [total_width * 0.35, total_width * 0.15, total_width * 0.15, total_width * 0.20,
                      total_width * 0.15]

        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 8, header, border=1, align='C', fill=True)
        pdf.ln()

        pdf.set_font('Roboto', '', 9)
        if not candidate.get('subjects'):
            pdf.cell(sum(col_widths), 10, "No subjects found for this candidate.", 1, new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align='C')
        else:
            for subject in candidate['subjects']:
                code = subject['code']
                cand_type = subject['type']
                subject_name = SUBJECT_CODE_MAP.get(code, code)
                papers_info = timetable.get(code, [])

                papers_to_show = papers_info
                if cand_type == 'R':
                    papers_to_show = [p for p in papers_info if p.get('paper') in ['1', '2']]

                for paper_info in papers_to_show:
                    pdf.cell(col_widths[0], 8, f" {subject_name}", 1)
                    pdf.cell(col_widths[1], 8, cand_type, 1, align='C')
                    pdf.cell(col_widths[2], 8, paper_info.get('paper', ''), 1, align='C')
                    pdf.cell(col_widths[3], 8, paper_info.get('date', ''), 1, align='C')
                    pdf.cell(col_widths[4], 8, paper_info.get('session', ''), 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                             align='C')

        bottom_text = (
            "Starting times for all centers within a territory are 09:00 hr. for the morning (9AM) session and 13:00 hr. for the afternoon (1PM) session. The Local Registrar reserves the right to arrange candidates for the administering of examinations."
        )
        pdf.ln(5)
        pdf.set_font('Roboto', 'I', 9)
        pdf.multi_cell(0, 5, bottom_text, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.output(filepath)
        return filepath

    except Exception as e:
        print(f"Failed to create PDF for {candidate.get('name', 'Unknown')}: {e}")
        return False


# ---------------- APP ----------------
class ESlipGeneratorApp:
    """The main application class for the CXC E-Slip Generator.

    This class builds and manages the main graphical user interface (GUI) of the
    application. It handles user interactions, file selections, and orchestrates
    the entire process of parsing data and generating e-slips.
    """

    def __init__(self, root):
        """Initializes the ESlipGeneratorApp.

        Args:
            root (tk.Tk): The root tkinter window for the application.
        """
        self.root = root
        self.root.title("CXC E-Slip Generator V2")
        self.root.geometry("800x700")

        self.log_queue = queue.Queue()
        self.file_paths = {"candidates": "", "centres": "", "csv": ""}
        self.centre_file_widgets = []  # To toggle disabled/normal state
        self.output_dir = ""
        self.exam_type = tk.StringVar(value="CSEC")

        self.month_options = ["January", "May - June"]
        self.exam_month = tk.StringVar(value="May - June")  # Default
        self.exam_year = tk.StringVar(value=str(datetime.now().year))
        self.centre_list_available = tk.BooleanVar(value=True)

        self._start_time = None

        wrap = ttk.Frame(root, padding=12)
        wrap.pack(fill="both", expand=True)

        exam_fr = ttk.LabelFrame(wrap, text="1. Exam Settings", padding=10)
        exam_fr.pack(fill="x")

        ttk.Label(exam_fr, text="Exam Type:").grid(row=0, column=0, sticky="w")
        self.exam_type_menu = ttk.OptionMenu(exam_fr, self.exam_type, self.exam_type.get(), "CSEC", "CAPE",
                                             command=self._update_month_options)
        self.exam_type_menu.grid(row=0, column=1, sticky="w")

        ttk.Label(exam_fr, text="Month:").grid(row=0, column=2, sticky="e", padx=(10, 0))
        self.month_menu = ttk.Combobox(exam_fr, textvariable=self.exam_month, width=12, state="readonly")
        self.month_menu.grid(row=0, column=3, sticky="w")

        ttk.Label(exam_fr, text="Year:").grid(row=0, column=4, sticky="e", padx=(10, 0))
        ttk.Entry(exam_fr, textvariable=self.exam_year, width=8).grid(row=0, column=5, sticky="w")
        exam_fr.grid_columnconfigure(6, weight=1)

        files_fr = ttk.LabelFrame(wrap, text="2. Select Source Files", padding=10)
        files_fr.pack(fill="x", pady=(10, 0))

        self._add_file_row(files_fr, 0, "Candidate List (PDF):", "candidates", [("PDF files", "*.pdf")])

        # Add Centre List row and store widgets
        self.centre_file_widgets = self._add_file_row(files_fr, 1, "Centre List (PDF):", "centres",
                                                      [("PDF files", "*.pdf")])

        # Add Checkbox for Centre List
        self.centre_list_checkbox = ttk.Checkbutton(files_fr, text="Centre List Available?",
                                                    variable=self.centre_list_available,
                                                    command=self._toggle_centre_list_input)
        self.centre_list_checkbox.grid(row=1, column=3, sticky="w", padx=(10, 0))

        self._add_file_row(files_fr, 2, "Eligibility CSV:", "csv", [("CSV files", "*.csv")])

        out_fr = ttk.LabelFrame(wrap, text="3. Output", padding=10)
        out_fr.pack(fill="x", pady=(10, 0))
        self.out_lbl = ttk.Label(out_fr, text="No folder selected", relief="sunken")
        self.out_lbl.grid(row=0, column=0, sticky="ew")
        ttk.Button(out_fr, text="Choose Folder", command=self.select_output_dir).grid(row=0, column=1, padx=6)
        out_fr.grid_columnconfigure(0, weight=1)

        act = ttk.Frame(wrap)
        act.pack(fill="x", pady=(10, 0))
        self.btn_start = ttk.Button(act, text="Generate E-Slips", command=self.start)
        self.btn_start.pack(side="left")

        self.progress_bar = ttk.Progressbar(act, orient="horizontal", mode="determinate")
        self.progress_bar.pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.status_label = ttk.Label(wrap, text="Status: Ready")
        self.status_label.pack(fill="x", pady=(5, 0))

        log_fr = ttk.LabelFrame(wrap, text="Log", padding=10)
        log_fr.pack(fill="both", expand=True, pady=(10, 0))

        log_container = ttk.Frame(log_fr)
        log_container.pack(fill="both", expand=True)

        self.log_txt = tk.Text(log_container, height=16, font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(log_container, orient="vertical", command=self.log_txt.yview)
        self.log_txt.configure(yscrollcommand=scrollbar.set)

        self.log_txt.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.root.after(100, self._drain_log)
        self.timetable_cache = {}

        # Initialize UI states
        self._update_month_options()
        self._toggle_centre_list_input()

    def _update_month_options(self, *args):
        """Updates the available exam months based on the selected exam type.

        This method is triggered when the user changes the "Exam Type" dropdown.
        It restricts the month options to "May - June" for CAPE exams, while
        allowing both "January" and "May - June" for CSEC exams.
        """
        exam_type = self.exam_type.get()
        if exam_type == "CAPE":
            self.month_menu['values'] = ["May - June"]
            self.exam_month.set("May - June")
        elif exam_type == "CSEC":
            self.month_menu['values'] = self.month_options  # ["January", "May - June"]

    def _toggle_centre_list_input(self, *args):
        """Enables or disables the centre list file input widgets.

        This method is connected to the "Centre List Available?" checkbox. It
        toggles the state of the file selection widgets for the centre list
        based on whether the checkbox is checked or not.
        """
        state = "normal" if self.centre_list_available.get() else "disabled"
        for widget in self.centre_file_widgets:
            widget.config(state=state)

        # Also clear the file path if disabled
        if not self.centre_list_available.get():
            self.file_paths["centres"] = ""
            # Find the label widget (it's the second one) and update its text
            if len(self.centre_file_widgets) > 1:
                self.centre_file_widgets[1].config(text="No file selected")

    def _add_file_row(self, parent, row, label, key, ftypes):
        """Creates a row of widgets for selecting a file.

        This helper method abstracts the creation of a labeled file selection
        row, which includes a label, a display for the selected file path, and a
        "Browse" button.

        Args:
            parent (tk.Widget): The parent widget for the new row.
            row (int): The grid row number to place the widgets in.
            label (str): The text label for the file input.
            key (str): The key to use for storing the file path.
            ftypes (list): A list of file type tuples for the file dialog.

        Returns:
            list: A list of the created widgets.
        """
        lbl_widget = ttk.Label(parent, text=label)
        lbl_widget.grid(row=row, column=0, sticky="w")

        path_lbl = ttk.Label(parent, text="No file selected", relief="sunken")
        path_lbl.grid(row=row, column=1, sticky="ew", padx=6)

        btn_widget = ttk.Button(parent, text="Browse",
                                command=lambda k=key, l=path_lbl, ft=ftypes: self._pick_file(k, l, ft))
        btn_widget.grid(row=row, column=2)

        parent.grid_columnconfigure(1, weight=1)

        # Return the widgets so they can be enabled/disabled
        return [lbl_widget, path_lbl, btn_widget]

    def _pick_file(self, key, label_widget, ftypes):
        """Opens a file dialog and updates the corresponding path.

        This method is called when the user clicks a "Browse" button. It opens
        a file dialog, and if a file is selected, it updates the internal
        `file_paths` dictionary and the label widget that displays the file
        name.

        Args:
            key (str): The key corresponding to the file type being selected.
            label_widget (tk.Widget): The label widget to update with the file
                                    name.
            ftypes (list): A list of file type tuples for the file dialog.
        """
        p = filedialog.askopenfilename(filetypes=ftypes)
        if p:
            self.file_paths[key] = p
            label_widget.config(text=os.path.basename(p))

    def select_output_dir(self):
        """Opens a directory dialog for selecting the output folder.

        This method is called when the user clicks the "Choose Folder" button.
        It opens a directory selection dialog and, if a directory is chosen,
        updates the output directory path and the corresponding label.
        """
        p = filedialog.askdirectory()
        if p:
            self.output_dir = p
            self.out_lbl.config(text=p)

    def log(self, msg):
        """Adds a message to the log queue to be displayed in the GUI.

        This method provides a thread-safe way to log messages from different
        parts of the application. Messages are added to a queue and then
        processed by the `_drain_log` method in the main GUI thread.

        Args:
            msg (str): The message to be logged.
        """
        self.log_queue.put(msg)

    def _drain_log(self):
        """Periodically checks the log queue and updates the log widget.

        This method is called repeatedly by the tkinter main loop to process
        any messages that have been added to the log queue. It updates the
        log text widget with new messages.
        """
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                self.log_txt.insert("end", msg + "\n")
                self.log_txt.see("end")
            except queue.Empty:
                break
        self.root.after(120, self._drain_log)

    def start(self):
        """Starts the e-slip generation process.

        This method is called when the "Generate E-Slips" button is clicked. It
        performs initial validation to ensure that all required files and the
        output directory have been selected. If validation passes, it starts a
        new thread to run the main processing logic.
        """
        # --- NEW VALIDATION LOGIC ---
        required_files = ["candidates", "csv"]
        if self.centre_list_available.get():
            required_files.append("centres")

        missing = [key for key in required_files if not self.file_paths[key]]

        if missing or not self.output_dir:
            msg = "Please select all required source files and an output folder.\n"
            if missing:
                msg += f"Missing files: {', '.join(missing)}"
            messagebox.showerror("Missing Files/Folder", msg)
            return

        if not self.exam_month.get().strip() or not self.exam_year.get().strip():
            messagebox.showerror("Missing Exam Details", "Please enter exam month and year.")
            return

        self._start_time = time.time()
        self.btn_start.config(state=tk.DISABLED)
        self.status_label.config(text="Status: Starting...")
        self.log_txt.delete("1.0", tk.END)

        t = threading.Thread(target=self._run)
        t.daemon = True
        t.start()

    def _run(self):
        """The main processing thread for generating e-slips.

        This method orchestrates the entire e-slip generation process, from
        parsing the input files to generating the final PDF outputs. It is run
        in a separate thread to keep the GUI responsive during processing.
        """
        try:
            exam_type = self.exam_type.get().strip()
            exam_month = self.exam_month.get().strip()
            exam_year = self.exam_year.get().strip()

            self.log("Status: Parsing CSV for eligible candidates...")
            # --- MODIFIED: Pass exam_type and exam_month to parser ---
            csv_list = parse_csv(self.file_paths["csv"], exam_type, exam_month, self.log)
            if not csv_list:
                self.log("No eligible candidates found in CSV. Stopping.")
                return

            # --- MODIFIED: Conditionally parse centre list ---
            centres = {}
            if self.centre_list_available.get():
                self.log("Status: Parsing Centre List...")
                centres = parse_centre_list(self.file_paths["centres"], self.log, self.output_dir)
                if not centres:
                    self.log("No centres found. Stopping.")
                    return
            else:
                self.log("Status: Skipping Centre List (not available).")

            self.log("Status: Parsing Candidate List...")
            # MODIFIED: Capture the full PDF text
            cand_list, unmatched_blocks, pdf_text = parse_candidate_list(self.file_paths["candidates"], self.log,
                                                                         self.output_dir)

            if unmatched_blocks:
                self.log(f"Opening manual candidate entry for {len(unmatched_blocks)} unmatched block(s)...")
                # MODIFIED: Pass pdf_text to the handler
                self.root.after(0, lambda: self._show_manual_candidate_entry(unmatched_blocks, cand_list, csv_list,
                                                                             centres, exam_month, exam_year, exam_type,
                                                                             pdf_text))
                return
            else:
                # MODIFIED: Pass pdf_text to continue processing
                self._continue_processing(cand_list, csv_list, centres, exam_month, exam_year, exam_type, pdf_text)

        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.root.after(0, self._reset_ui)

    def _show_manual_candidate_entry(self, unmatched_blocks, cand_list, csv_list, centres, exam_month, exam_year,
                                     exam_type, pdf_text):
        """Displays the manual candidate entry dialog and handles the results.

        This method is called when there are unparsable blocks of text from the
        candidate list PDF. It opens a dialog that allows the user to manually
        enter the details for these candidates.

        Args:
            unmatched_blocks (list): A list of unparsable text blocks.
            cand_list (list): The list of successfully parsed candidates.
            csv_list (list): The list of eligible candidates from the CSV.
            centres (dict): A dictionary of exam centres.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
            pdf_text (str): The full text from the candidate list PDF.
        """
        # MODIFIED: Pass pdf_text to the dialog
        dlg = ManualCandidateEntry(self.root, unmatched_blocks, pdf_text)
        extra = dlg.show()
        if extra:
            self.log(f"Added {len(extra)} candidates from manual entry")
            cand_list.extend(extra)

        t = threading.Thread(
            target=lambda: self._continue_processing(cand_list, csv_list, centres, exam_month, exam_year, exam_type,
                                                     pdf_text))
        t.daemon = True
        t.start()

    def _continue_processing(self, cand_list, csv_list, centres, exam_month, exam_year, exam_type, pdf_text):
        """Continues the processing after the initial PDF parsing.

        This method is responsible for cross-matching the candidates parsed from
        the PDF with the eligible candidates from the CSV. It identifies which
        candidates are matched and which are missing.

        Args:
            cand_list (list): The list of candidates parsed from the PDF.
            csv_list (list): The list of eligible candidates from the CSV.
            centres (dict): A dictionary of exam centres.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
            pdf_text (str): The full text from the candidate list PDF.
        """
        try:
            self.log("Status: Cross-matching candidates with CSV...")
            matched = []
            missing_csv = []

            def make_key(c):
                name_key = normalize_key_name(c.get('name', ''))
                return (name_key, c.get('dob', ''))

            self.log("\n--- DEBUG: Building searchable index from PDF data... ---")
            cand_index = {}
            for c in cand_list:
                key = make_key(c)
                if key[0]:
                    cand_index[key] = c
                    self.log(f"  -> Indexing PDF Key: {key} -> Value: {c['name']}")

            self.log(f"\n--- DEBUG: Built PDF index with {len(cand_index)} unique candidate keys. ---")
            self.log("--- DEBUG: Now comparing each CSV entry against this index. ---")

            for row in csv_list:
                csv_name_key = normalize_key_name(row.get('name', ''))
                k = (csv_name_key, row.get('dob', ''))
                self.log(f"Checking CSV Key: {k}")

                if k in cand_index:
                    matched.append(cand_index[k])
                    self.log(f"  -> MATCH FOUND!")
                    self.log(f"  -> PDF Index Value: {cand_index[k]}")
                else:
                    missing_csv.append(row)
                    self.log("  -> !! MATCH NOT FOUND in PDF index. !!")

            self.log(f"\nMatched {len(matched)} candidates")

            if missing_csv:
                self.log(f"CSV candidates not found in candidate list: {len(missing_csv)}. Opening manual entry...")
                # MODIFIED: Pass pdf_text to the next step
                self.root.after(0, lambda: self._show_manual_csv_entry(missing_csv, matched, centres, exam_month,
                                                                       exam_year, exam_type, pdf_text))
                return
            else:
                self._continue_with_centres(matched, centres, exam_month, exam_year, exam_type)

        except Exception as e:
            self.log(f"ERROR in continue processing: {e}")
            self.root.after(0, self._reset_ui)

    def _show_manual_csv_entry(self, missing_csv, matched, centres, exam_month, exam_year, exam_type, pdf_text):
        """Displays the manual CSV entry dialog for unmatched candidates.

        This method is called when there are candidates in the CSV file who
        could not be matched with any candidates from the PDF. It opens a
        dialog that allows the user to manually enter the candidate IDs for
        these unmatched candidates.

        Args:
            missing_csv (list): A list of unmatched candidates from the CSV.
            matched (list): The list of successfully matched candidates.
            centres (dict): A dictionary of exam centres.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
            pdf_text (str): The full text from the candidate list PDF.
        """
        # MODIFIED: Pass pdf_text to the dialog
        dlg = ManualCSVEntry(self.root, missing_csv, pdf_text)
        extra = dlg.show()
        if extra:
            self.log(f"Added {len(extra)} candidates from CSV manual entry")
            matched.extend(extra)

        t = threading.Thread(
            target=lambda: self._continue_with_centres(matched, centres, exam_month, exam_year, exam_type))
        t.daemon = True
        t.start()

    def _continue_with_centres(self, matched, centres, exam_month, exam_year, exam_type):
        """Continues the processing after matching candidates, handling centres.

        This method checks for any missing centre names. If the centre list is
        available, it identifies any centre codes from the matched candidates
        that do not have a corresponding name in the `centres` dictionary and
        prompts the user for manual entry if needed.

        Args:
            matched (list): The list of matched candidates.
            centres (dict): A dictionary of exam centres.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
        """
        # --- MODIFIED: Conditionally run this entire step ---
        if self.centre_list_available.get():
            try:
                needed_centres = {c['centre_num'] for c in matched if c.get('centre_num')}
                missing_codes = sorted([c for c in needed_centres if c not in centres])

                if missing_codes:
                    self.log(f"Missing {len(missing_codes)} centre code(s). Opening manual centre entry...")
                    self.root.after(0,
                                    lambda: self._show_manual_centre_entry(missing_codes, centres, matched, exam_month,
                                                                           exam_year, exam_type))
                    return
                else:
                    self._continue_with_timetable(matched, centres, exam_month, exam_year, exam_type)

            except Exception as e:
                self.log(f"ERROR in centre processing: {e}")
                self.root.after(0, self._reset_ui)
        else:
            # --- NEW: Skip centre check if not available ---
            self.log("Skipping manual centre entry (not available).")
            # Pass the empty 'centres' dict straight to the timetable step
            self._continue_with_timetable(matched, centres, exam_month, exam_year, exam_type)

    def _show_manual_centre_entry(self, missing_codes, centres, matched, exam_month, exam_year, exam_type):
        """Displays the manual centre entry dialog.

        This method is called when there are missing centre names. It opens a
        dialog that allows the user to manually enter the names for the missing
        centre codes.

        Args:
            missing_codes (list): A list of centre codes without names.
            centres (dict): The dictionary of existing centre mappings.
            matched (list): The list of matched candidates.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
        """
        dlg = ManualCentreEntry(self.root, missing_codes)
        added = dlg.show()
        if added:
            centres.update(added)
            self.log(f"Added {len(added)} centre mappings")

        t = threading.Thread(
            target=lambda: self._continue_with_timetable(matched, centres, exam_month, exam_year, exam_type))
        t.daemon = True
        t.start()

    def _continue_with_timetable(self, matched, centres, exam_month, exam_year, exam_type):
        """Continues the processing after handling centres, moving on to the timetable.

        This method gathers all unique subjects from the list of matched
        candidates and then prompts the user to enter the timetable information
        for those subjects.

        Args:
            matched (list): The list of matched candidates.
            centres (dict): A dictionary of exam centres.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
        """
        try:
            subject_universe = set()
            for c in matched:
                for s in c.get('subjects', []):
                    subject_universe.add(s['code'])
            subject_universe = {s for s in subject_universe if s}

            if ASK_TIMETABLE_EVERY_RUN and subject_universe:
                self.log(f"Collecting timetable for {len(subject_universe)} unique subject(s)...")
                self.root.after(0, lambda: self._show_manual_timetable_entry(subject_universe, matched, centres,
                                                                             exam_month, exam_year, exam_type))
                return
            else:
                self._generate_slips(matched, centres, {}, exam_month, exam_year, exam_type)

        except Exception as e:
            self.log(f"ERROR in timetable processing: {e}")
            self.root.after(0, self._reset_ui)

    def _show_manual_timetable_entry(self, subject_universe, matched, centres, exam_month, exam_year, exam_type):
        """Displays the manual timetable entry dialog.

        This method opens a dialog that allows the user to enter the timetable
        details for all the unique subjects found among the matched candidates.

        Args:
            subject_universe (set): A set of all unique subject codes.
            matched (list): The list of matched candidates.
            centres (dict): A dictionary of exam centres.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
        """
        dlg = ManualTimetableEntry(self.root, subject_universe, exam_month, exam_year)
        tt = dlg.show()
        if tt:
            self.timetable_cache.update(tt)
            self.log(f"Added timetable for {len(tt)} subjects")

        t = threading.Thread(
            target=lambda: self._generate_slips(matched, centres, self.timetable_cache, exam_month, exam_year,
                                                exam_type))
        t.daemon = True
        t.start()

    def _generate_slips(self, matched, centres, timetable, exam_month, exam_year, exam_type):
        """Generates the final PDF e-slips for all matched candidates.

        This is the final step in the process. It iterates through all the
        matched candidates and calls the `create_pdf_slip` function to generate
        a PDF e-slip for each one.

        Args:
            matched (list): The final list of matched and verified candidates.
            centres (dict): A dictionary of exam centres.
            timetable (dict): A dictionary of the exam timetable.
            exam_month (str): The month of the exam.
            exam_year (str): The year of the exam.
            exam_type (str): The type of the exam.
        """
        try:
            self.log("Status: Generating PDF slips...")
            total_candidates = len(matched)
            self.progress_bar["maximum"] = total_candidates
            success_count = 0

            for i, c in enumerate(matched):
                self.progress_bar["value"] = i + 1

                centre_name = centres.get(c.get('centre_num', ''), '')

                try:
                    out = create_pdf_slip(c, centre_name, timetable, self.output_dir, exam_month, exam_year, exam_type)
                    if out:
                        # --- FIX 2: Changed os.basename to os.path.basename ---
                        self.log(f"Generated: {os.path.basename(out)}")
                        success_count += 1
                    else:
                        self.log(f"Failed to generate slip for {c.get('name', 'Unknown')}")
                except Exception as e:
                    self.log(f"Failed to generate slip for {c.get('name', 'Unknown')}: {e}")

            duration = time.time() - (self._start_time or time.time())
            final_message = f"Complete! {success_count}/{total_candidates} slips generated in {duration:.2f}s."
            self.log(final_message)
            self.status_label.config(text=f"Status: {final_message}")

            messagebox.showinfo("Success",
                                f"Processing complete.\n{success_count} of {total_candidates} e-slips were generated.")

        except Exception as e:
            self.log(f"ERROR in slip generation: {e}")
            messagebox.showerror("Error", f"An error occurred during slip generation: {e}")
        finally:
            self.root.after(0, self._reset_ui)

    def _reset_ui(self):
        """Resets the UI to its initial state after processing is complete.

        This method is called at the end of the e-slip generation process,
        whether it succeeds or fails. It re-enables the "Generate E-Slips"
        button and resets the progress bar.
        """
        self.btn_start.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0


# ---------------- MAIN ----------------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("CXC E-Slip Generator")
    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except Exception:
        pass
    app = ESlipGeneratorApp(root)
    root.mainloop()