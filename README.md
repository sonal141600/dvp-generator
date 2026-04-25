# ⚡ SpecSync — DVP Test Plan Generator

> Automatically generate Design Validation Plan (DVP) Test Plans from automotive RFQ engineering drawings using AI.

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Flask](https://img.shields.io/badge/Flask-3.0-lightgrey)
![Claude API](https://img.shields.io/badge/Claude-Sonnet-orange)
![Google Vision](https://img.shields.io/badge/Google-Vision%20API-4285F4)

---

## 🧠 What is SpecSync?

In the automotive supplier industry, every new part requires a **DVP (Design Validation Plan)** — a document listing all the tests the part must pass, which standards apply, and whether the test specifications are available.

Creating DVPs manually is time-consuming. Engineers read engineering drawings, identify test standards (like `TL 1010`, `DIN EN ISO 845`, `TSL3608G`), check which spec documents are available, and fill in acceptance criteria — all by hand.

**SpecSync automates this entire process.**

Upload a zipped RFQ folder containing engineering drawings and spec PDFs. SpecSync reads the drawing using OCR + AI, extracts all test standards, checks their availability against your standards library, fills in acceptance criteria from spec PDFs, and outputs a formatted Excel DVP Test Plan — in minutes.

---

## 🎯 Key Features

- **Multi-customer support** — Works with VW, Toyota, Suzuki, Skoda and more
- **Smart PDF classifier** — Automatically sorts drawings, spec docs, and irrelevant files
- **Google Vision OCR** — Accurate text extraction from TIF and image-based PDF drawings
- **Claude AI understanding** — Converts raw OCR text into structured JSON test data
- **Criteria auto-extraction** — Reads spec PDFs and extracts key acceptance criteria in one line
- **Standards library merge** — Combines your uploaded standards with the server library
- **Availability checking** — Marks each test AVAILABLE (green) or NOT AVAILABLE (red)
- **Excel output** — Professional formatted DVP Test Plan ready to send to customers
- **Web interface** — Clean drag-and-drop UI built with Flask

---

## 🛠 Tech Stack

| Component | Technology |
|---|---|
| Backend | Python, Flask |
| AI Understanding | Anthropic Claude Sonnet |
| OCR | Google Cloud Vision API |
| Drawing Processing | PIL (Pillow), pdf2image |
| PDF Reading | pdfplumber |
| Excel Output | openpyxl |
| Frontend | HTML, CSS, JavaScript (vanilla) |

---

## 🔄 How It Works

```
RFQ Folder (zip)
      │
      ▼
1. PDF Classifier
   ├── Drawing (TIF / image PDF)
   ├── Spec docs (TSL, VW, DIN PDFs)
   └── Skip (CPL, quotation, SOR etc.)
      │
      ▼
2. Drawing Reader (3×3 or 4×3 grid scan)
   ├── Google Vision OCR → extracts raw text
   └── Claude AI → extracts standards + criteria as JSON
      │
      ▼
3. Standards Library Check
   ├── Check against server library
   ├── Merge with user-uploaded standards
   └── Mark AVAILABLE / NOT AVAILABLE
      │
      ▼
4. Criteria Extraction
   └── Read spec PDFs → extract key criteria in one line
      │
      ▼
5. Excel DVP Output
   └── Formatted test plan with availability + criteria
```

---

## 🚀 Setup

### Prerequisites

- Python 3.10+
- Anthropic API key — [console.anthropic.com](https://console.anthropic.com)
- Google Cloud Vision API key — [console.cloud.google.com](https://console.cloud.google.com)

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/specsync.git
cd specsync

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Mac/Linux
# venv\Scripts\activate   # Windows

# Install dependencies
pip install -r requirements.txt
```

### Configuration

Create a `.env` file in the project root:

```
ANTHROPIC_API_KEY=your_anthropic_api_key_here
GOOGLE_APPLICATION_CREDENTIALS=rfq_vision_key.json
```

Place your Google Cloud Vision service account JSON file in the project root as `rfq_vision_key.json`.

### Standards Library

Create a `standards_library/` folder and add your spec PDFs:

```
standards_library/
├── TL1010_Materials for Vehicle Interiors.pdf
├── VW_50026_EN.pdf
├── VW_50180_EN.pdf
└── ...
```

---

## 💻 Usage

### Web Interface (recommended)

```bash
python app.py
```

Open `http://127.0.0.1:5000` in your browser.

1. Upload your RFQ folder as a `.zip` file
2. Enter your company name (optional)
3. Upload your standards library as a `.zip` (optional — merged with server library)
4. Click **Generate DVP Test Plan**
5. Download the Excel output

### Command Line

```bash
# Process a single RFQ folder
python dvp_reader.py rfq_inputs/RFQ_Toyota_002

# Process all folders in rfq_inputs/
python main.py
```

---

## 📁 Project Structure

```
specsync/
├── app.py              # Flask web application
├── dvp_reader.py       # Core DVP generation logic
├── main.py             # CLI batch processor
├── templates/
│   └── index.html      # Web UI
├── standards_library/  # Your spec PDFs (not committed)
├── output/             # Generated DVP files (not committed)
├── requirements.txt
├── .env.example
└── .gitignore
```

---

## 📊 Output

The generated Excel DVP Test Plan includes:

| Column | Description |
|---|---|
| Serial No. | Test sequence number |
| Test Description | Name of the test |
| Test Method | Standard code (e.g. TL 1010, DIN EN ISO 845) |
| Responsibility | Who performs the test |
| Test Criteria | Acceptance criteria extracted from spec |
| Test Start/End Date | To be filled by engineer |
| Test Agency | Lab performing the test |
| Remarks | AVAILABLE (green) / NOT AVAILABLE (red) |

---

## ⚙️ Supported Customers

| Customer | Drawing Format | Notes |
|---|---|---|
| Volkswagen / Skoda | TIF | German + English notes |
| Toyota | Image PDF | Japanese + English, TSL/TSM/TSZ standards |
| Suzuki | TIF | SES standards in notes |
| Generic | TIF / PDF | Any standard format |

---

## 🔑 Environment Variables

| Variable | Description |
|---|---|
| `ANTHROPIC_API_KEY` | Your Anthropic API key |
| `GOOGLE_APPLICATION_CREDENTIALS` | Path to Google Vision service account JSON |

---

## 📝 License

MIT License — feel free to use and modify.

---

## 👤 Author

Built as a portfolio project to demonstrate real-world AI automation in the automotive supplier industry.
=======
# dvp-generator
AI-powered DVP Test Plan generator for automotive suppliers — reads engineering drawings via OCR, extracts test standards, checks availability, and outputs formatted Excel reports.
>>>>>>> 4c7ece61c20ece0029130416807936e654c0560e
