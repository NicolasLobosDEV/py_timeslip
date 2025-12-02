### Instructions to Build

1.  **Organize your folder:**
    Put all these files in the **same folder** as `ESlipGenerator.py`:
    * `background.png`
    * `Roboto-Regular.ttf`
    * `Roboto-Bold.ttf`
    * `Roboto-Italic.ttf`
    * **Folder:** `Tesseract-OCR` (Copy this from `C:\Program Files\Tesseract-OCR`)

2.  **Run this PyInstaller command:**

    ```bash
    pyinstaller --noconsole --onefile --add-data "Tesseract-OCR;Tesseract-OCR" --add-data "Roboto-Regular.ttf;." --add-data "Roboto-Bold.ttf;." --add-data "Roboto-Italic.ttf;." --add-data "background.png;." timeslips.py