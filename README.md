# CXC E-Slip Generator

This application automates the process of creating and distributing personalized exam e-slips for students participating in CXC examinations. It is designed to parse candidate and centre information from PDF and CSV files, cross-reference the data, and generate individual PDF e-slips with customized timetables.

## Features

- **Automated Parsing:** Extracts candidate and exam centre data from PDF and CSV files, reducing manual data entry.
- **Intelligent Matching:** Cross-references data from multiple sources to ensure accuracy and identify any discrepancies.
- **Manual Correction Tools:** Provides a user-friendly interface for manual data entry and correction when automated parsing is not possible.
- **Customizable Timetables:** Allows for manual entry of exam timetables to ensure that each e-slip contains the correct information.
- **PDF Generation:** Generates high-quality, professional-looking PDF e-slips for each candidate.

## Setup and Installation

### Prerequisites

- Python 3.x
- Tesseract OCR Engine

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-repo/py_timeslip.git
   ```

2. **Install the required Python libraries:**
   ```bash
   pip install PyMuPDF fpdf2 pytesseract Pillow
   ```

3. **Install Tesseract:**
   - **Windows:** Download and run the installer from the official [Tesseract repository](https://github.com/tesseract-ocr/tesseract).
   - **macOS:** Install using Homebrew:
     ```bash
     brew install tesseract
     ```
   - **Linux:** Install using your distribution's package manager:
     ```bash
     sudo apt-get install tesseract-ocr
     ```

## Usage

1. **Launch the application:**
   ```bash
   python timeslips.py
   ```

2. **Select the source files:**
   - **Candidate List (PDF):** The PDF file containing the list of all candidates.
   - **Centre List (PDF):** The PDF file listing all exam centres and their codes.
   - **Eligibility CSV:** The CSV file containing the list of candidates who are eligible to receive an e-slip.

3. **Choose an output folder:**
   - Select the directory where the generated PDF e-slips will be saved.

4. **Generate the e-slips:**
   - Click the "Generate E-Slips" button to start the process.

5. **Manual Entry (if required):**
   - If the application encounters any data that it cannot parse automatically, it will open a dialog window for you to enter the information manually.
