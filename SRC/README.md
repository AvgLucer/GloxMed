# GloxMed - Source

This folder contains the main source code for **GloxMed**, an AI-powered medical document explainer developed by **Glox Industries**.

## 📁 Source

```text
GloxMed/
└── SRC/
    ├── gloxmed.py
    └── README.md
```

### `gloxmed.py`

The main GloxMed application.

It contains:

* PySide6 desktop UI
* PDF document extraction
* DOCX document extraction
* Patient profile management
* Doctor instructions and symptom input
* OpenRouter AI integration
* Free-model discovery
* Structured medical explanation generation
* ReportLab PDF report generation
* Local report storage
* Local `.env` configuration
* Background AI processing using `QThread`

## 🛠️ Requirements

Python **3.10+** is recommended.

Install the required packages:

```bash
pip install PySide6 python-dotenv openrouter PyMuPDF python-docx reportlab
```

## 🔑 OpenRouter API Key

GloxMed uses an OpenRouter API key for AI generation.

The application automatically looks for:

```text
.env
```

When running from source, it is expected beside `gloxmed.py`:

```text
GloxMed/
└── SRC/
    ├── gloxmed.py
    ├── .env
    └── README.md
```

Add:

```env
OPENROUTER_API_KEY=your_api_key_here
```

GloxMed can also create the `.env` file automatically if it does not exist.

## ▶️ Run From Source

From the `SRC` directory:

```bash
python gloxmed.py
```

Or:

```bash
py gloxmed.py
```

## 📄 Supported Documents

GloxMed currently supports:

* `.pdf`
* `.docx`

PDFs must contain selectable text. Image-only or scanned PDFs require OCR before their contents can be extracted.

## 💾 Local Data

When running from source, GloxMed stores its local application data alongside the source file.

```text
SRC/
├── gloxmed.py
├── requirements.txt
|── README.md
```

Generated reports are stored inside:

```text
GloxMed_Reports/
```

The application does not require a separate database for patient profiles or generated reports.

## 📦 Building the EXE

GloxMed can be packaged using PyInstaller.

Install PyInstaller:

```bash
pip install pyinstaller
```

Build:

```bash
pyinstaller --onefile --windowed --name GloxMed --collect-all PySide6 --collect-all openrouter --collect-all pymupdf gloxmed.py
```

The executable will be generated in:

```text
dist/
└── GloxMed.exe
```

### EXE Data Location

When running as a PyInstaller executable, GloxMed uses the **folder containing `GloxMed.exe`** as its application directory.

For example:

```text
GloxMed/
└── dist/
    ├── GloxMed.exe
    ├── .env
    ├── gloxmed_profiles.json
    └── GloxMed_Reports/
```

The `.env` file is therefore **not embedded inside the executable**.

If it does not exist beside the executable, GloxMed creates it automatically.

## ⚠️ Medical Safety

GloxMed is an **educational information tool**, not a medical diagnostic system.

The generated explanations are intended to help users understand medical documents and prepare questions for a qualified healthcare professional.

GloxMed does not intentionally:

* Diagnose medical conditions
* Prescribe medication
* Recommend medication dosage changes
* Replace a doctor
* Invent medical findings or values

Always consult a qualified healthcare professional for medical decisions.

## 🔒 Privacy

Medical documents are processed locally for text extraction and report generation.

However, when AI explanation is requested, the extracted document information and user-provided context are sent to the selected **OpenRouter model** for processing.

Do not upload sensitive medical information unless you understand and accept the privacy policies of the AI service and model being used.

## 🏗️ Technology Stack

| Technology    | Purpose                      |
| ------------- | ---------------------------- |
| Python        | Core application             |
| PySide6       | Desktop interface            |
| OpenRouter    | AI model access              |
| python-dotenv | Local API key configuration  |
| PyMuPDF       | PDF text extraction          |
| python-docx   | DOCX text extraction         |
| ReportLab     | PDF report generation        |
| PyInstaller   | Windows executable packaging |

## © Glox Industries

**GloxMed** is developed by **Glox Industries**.

> Building Softwares With a New Vision

Source code is provided for educational and development purposes.
