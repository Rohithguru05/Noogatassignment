# AI-Powered PowerPoint Inconsistency Analyzer

This command-line tool leverages Google's Gemini AI to perform a deep analysis of PowerPoint (`.pptx`) presentations. It intelligently identifies factual, numerical, and logical inconsistencies across slides by analyzing both native text and text embedded within images.

## How It Works

The tool operates in a simple three-stage pipeline:

1.  **Extract:** It iterates through every slide, extracting all standard text from titles, shapes, and notes. Crucially, it uses an advanced Optical Character Recognition (OCR) engine to "read" and extract text from any images, charts, or diagrams.
2.  **Analyze:** The consolidated text from all slides is sent to the Gemini AI model. The AI is prompted with a specific set of instructions to act as an expert business analyst, looking for contradictions, mathematical errors, and logical gaps.
3.  **Report:** The AI's findings are received and formatted into a highly structured, colorful, and easy-to-read report directly in the terminal, with each issue presented in its own distinct box for maximum clarity.

## Key Features

* **Multi-Modal Text Extraction:** Reads text from virtually anywhere in your presentation, including text boxes, speaker notes, and, most importantly, images.
* **Intelligent Inconsistency Detection:** Goes beyond simple text matching to find a wide range of issues:
    * **Numerical Conflicts:** (e.g., "$2M revenue" on one slide vs. "$2.5M revenue" on another).
    * **Contradictory Claims:** (e.g., "We are the market leader" vs. "We are aiming for the #2 spot").
    * **Timeline Mismatches:** (e.g., "Launch in Q3" vs. a roadmap showing a Q4 launch).
    * **Missing Data:** Flags placeholders like "Source: TBD".
* **Advanced Usability:**
    * **Secure API Key Handling:** Uses a `config.ini` file, which is ignored by Git, to keep your secret API key safe and off of GitHub. An example template is provided.
    * **Intelligent Caching:** Automatically saves the results of an analysis. If you run the tool on the same file again, the report is delivered instantly, saving time and API costs.
    * **Multiple Output Formats:** Save the final report as a `.txt` or `.md` (Markdown) file for easy sharing.
* **Clarity of Output:** The terminal report is designed for maximum readability, using a colorful, boxed layout to clearly separate and explain each issue.

## Output Example

```
╔════════════════════════════════════════════════════════════════════════════════════╗
║                      AI Inconsistency Analysis Report                      ║
║                      Found 5 potential issue(s) to review.                     ║
╚════════════════════════════════════════════════════════════════════════════════════╝

╔═══════════════════════ ISSUE #1: Numerical Inconsistency ════════════════════════╗
║                                                                                    ║
║ CONFLICT:    Slide 1 states $2M impact, while Slide 2 states $3M saved annually  ║
║              in lost productivity hours.                                         ║
║                                                                                    ║
╠════════════════════════════════════════════════════════════════════════════════════╣
║ EVIDENCE:                                                                          ║
║   - Slide 1: '$2M Impact'                                                          ║
║   - Slide 2: '$3M saved in lost productivity hours annually'                       ║
╚════════════════════════════════════════════════════════════════════════════════════╝
```

## Setup and Installation (Windows)

### 1. Install the Tesseract OCR Engine

This is a one-time setup that is essential for reading text from images.

* **Download:** Get the installer from the official [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) page.
* **Install:** Run the installer. It is recommended to use the default installation path (usually `C:\Program Files\Tesseract-OCR`).
* **Add to System PATH:** This is the most critical step.
    * In the Windows Start Menu, search for `env` and select **"Edit the system environment variables"**.
    * Click **"Environment Variables..."**, then in the "System variables" section, double-click on **`Path`**.
    * Click **"New"** and paste the path to your Tesseract folder: `C:\Program Files\Tesseract-OCR`.
    * Click **OK** on all windows to save.
    * **You must restart any open terminal or VS Code window** for the change to take effect.

### 2. Set Up the Python Project

* **Create a Virtual Environment:**
    ```powershell
    python -m venv venv
    .\venv\Scripts\activate
    ```
* **Install Dependencies:**
    ```powershell
    pip install -r requirements.txt
    ```
* **Configure Your API Key:**
    1.  **Create Your Config File:** In the project folder, find the template file named `config.example.ini`. Make a **copy** of this file and rename the copy to `config.ini`.
    2.  **Add Your Key:** Open your new `config.ini` file. Paste your secret Google AI API key into the `api_key` field, replacing `YOUR_API_KEY_HERE`.

## How to Run

Activate your virtual environment (`.\venv\Scripts\activate`) and use one of the following commands:

* **Analyze a specific file:**
    ```powershell
    python main.py --file "path\to\your\presentation.pptx"
    ```
* **Analyze and save the report:**
    ```powershell
    python main.py --file "MyDeck.pptx" --output report.md
    ```
* **Force a re-analysis (ignore the cache):**
    ```powershell
    python main.py --file "MyDeck.pptx" --no-cache
    ```

## Limitations

* **OCR Accuracy:** The quality of analysis on image-based content is directly dependent on the clarity and resolution of the images. The OCR may struggle with blurry, handwritten, or highly stylized text.
* **Complex Inference:** The AI is extremely capable but may not understand deep, domain-specific context that is not explicitly mentioned in the text. It analyzes the information provided, but it does not have external knowledge of your specific project or company.
* **Read-Only:** This tool is for analysis only and will never make any modifications to your original `.pptx` file.
