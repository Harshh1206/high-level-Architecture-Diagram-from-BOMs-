# ⚡ DeltaV Architecture Generator

**BOM → Architecture Diagram in Seconds**

A rule-based automation tool that transforms Bill of Materials (BOM) documents into professional DeltaV system architecture diagrams using real hardware component images.

---

## 🎯 Overview

This application automates the tedious process of creating system architecture diagrams by:

1. **Parsing** any Bill of Materials file (Excel, CSV)
2. **Classifying** each component automatically using configurable rules
3. **Grouping** items into logical cabinets and areas
4. **Rendering** a professional architecture diagram with component images
5. **Exporting** as PowerPoint (.pptx) for easy integration into documentation

**Key Features:**
- ✅ Auto-detects column order in BOMs (Description, Qty, Area, Part No)
- ✅ Handles merged cells, section labels, and flexible headers
- ✅ Multi-encoding support (UTF-8, Latin-1, Windows-1252)
- ✅ Real-time classification feedback with confidence scores
- ✅ Self-learning: Save corrections to refine future classifications
- ✅ 100% offline operation with no external API calls
- ✅ Instant PNG preview of generated diagrams
- ✅ High-resolution PPTX export with component images

---

## 🚀 Quick Start

### Installation

1. **Clone/Download** the project to your machine
2. **Install Python 3.10+** (if not already installed)
3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

### Running the App

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

---

## 📋 How to Use

### Step 1: Upload BOM
- Click the upload area and select your BOM file (.xlsx, .xls, or .csv)
- The app auto-detects column positions and handles flexible formats

### Step 2: Review Classifications
- View the parsed table with auto-assigned component classes
- Use dropdown menus to manually correct any "UNKNOWN" classifications
- Save corrections to `rules.json` to auto-classify similar items in future BOMs

### Step 3: Download Architecture
- The diagram auto-generates as you make corrections
- Preview shows the PNG rendering
- Download the `.pptx` file for full resolution and editing in PowerPoint

---

## 📁 Project Structure

```
Image_BOM/
├── app.py                    # Streamlit web interface
├── parser.py                 # BOM file parsing engine
├── classifier.py             # Rule-based component classification
├── grouper.py                # Groups items into cabinets
├── generator.py              # PPTX + PNG rendering
├── Make_placeholders.py       # Utility for creating placeholder images
├── rules.json                # Classification rules (auto-updated)
├── handmade_demo.json        # Example BOM data
├── requirements.txt          # Python dependencies
├── images/                   # Component images (PNG/JPG)
└── README.md                 # This file
```

---

## 🔧 Architecture

### Pipeline Flow

```
BOM File → Parser → Classifier → Grouper → Generator → PPTX/PNG
            ↓          ↓            ↓         ↓
         Parse    Apply rules   Group by   Render
         CSV/XLS  (confidence)  cabinet    with images
```

### Key Modules

| Module | Purpose |
|--------|---------|
| **parser.py** | Extracts data from BOM files; handles flexible column orders, merged cells, and multiple encodings |
| **classifier.py** | Matches component descriptions against rule set; returns classification + confidence score |
| **grouper.py** | Organizes items into logical cabinets and areas (PDC Room, Operator Room) |
| **generator.py** | Creates PowerPoint slides with boxes, text, and component images; exports as PPTX and PNG |
| **app.py** | Streamlit UI; orchestrates the pipeline and manages user corrections |

---

## 📊 Supported BOM Formats

- **Excel**: `.xlsx` (Office Open XML), `.xls` (Legacy)
- **CSV**: Any delimiter (auto-detected)
- **Encodings**: UTF-8, Latin-1, Windows-1252
- **Headers**: Auto-detected on rows 1–15
- **Multi-sheet**: Best sheet auto-selected

### Expected BOM Columns

- **Description** — Component name/model
- **Qty** — Quantity (numeric)
- **Area** — Location (e.g., "PDC ROOM", "OPERATOR ROOM")
- **Part No** — Part number (optional)

Any column order works; the parser adapts automatically.

---

## 🧠 Classification System

### How Classification Works

1. **Rule Matching**: Component descriptions are matched against patterns in `rules.json`
2. **Confidence Scores**: Shown as percentages (e.g., 95% confidence)
3. **Classes**: Components are categorized as:
   - `PLC` — Programmable Logic Controller
   - `HMI` — Human Machine Interface
   - `IO_MODULE` — Input/Output Modules
   - `POWER_SUPPLY` — Power supplies
   - `CABINET` — Enclosures
   - `WORKSTATION` — Generic hardware
   - `UNKNOWN` — Unmatched (requires manual review)

### Self-Learning

After reviewing and correcting classifications:
1. Click **"Save corrections to rules.json"**
2. The app learns the new pattern
3. Future BOMs with similar descriptions auto-classify correctly

---

## 🎨 Component Images

The `images/` directory contains hardware component images (PNG/JPG) used in diagram rendering.

To add new components:
1. Place `.png` or `.jpg` files in the `images/` folder
2. Name them descriptively (e.g., `plc_s7_1200.png`)
3. They're automatically included in future renders

**Rendering Fallback**: If image rendering fails, the app shows a placeholder and allows PPTX download for manual viewing.

---

## 📦 Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| streamlit | Latest | Web UI framework |
| pandas | 2.2.3+ | BOM data handling |
| python-pptx | 1.0.2+ | PowerPoint generation |
| Pillow | 10.0.0+ | Image rendering |
| openpyxl | 3.1.5+ | Excel file support |
| xlrd | 2.0.1+ | Legacy XLS support |

See `requirements.txt` for the complete list.

---

## ⚙️ System Requirements

- **OS**: Windows, macOS, Linux
- **Python**: 3.10 or higher
- **RAM**: 512 MB minimum (1 GB recommended)
- **Disk**: 100 MB for app + dependencies
- **Optional**: LibreOffice + pdftoppm (for high-quality PDF→PNG conversion)

### For Enhanced Rendering (Linux/macOS)

Install LibreOffice for better diagram previews:

```bash
# macOS (Homebrew)
brew install libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice imagemagick

# CentOS/RHEL
sudo yum install libreoffice imagemagick
```

---

## 🐛 Troubleshooting

### "Preview unavailable" or PNG rendering fails
- **Cause**: Missing component images or rendering library
- **Fix**: Download component images to `images/` folder or install LibreOffice (see above)

### "use_container_width" parameter error
- **Cause**: Outdated Streamlit version
- **Fix**: `pip install --upgrade streamlit`

### BOM not parsing correctly
- **Cause**: Unusual encoding or structure
- **Fix**: Try converting to `.csv` first using Excel's "Save As" feature

### Classification accuracy low
- **Cause**: Rules outdated or incomplete
- **Fix**: Manually classify items → click "Save corrections" to retrain

---

## 🔐 Security & Privacy

- **100% Offline**: No data is sent to external servers
- **Local Processing**: All parsing, classification, and rendering happens locally
- **No Tracking**: No telemetry or usage analytics
- **Rules Storage**: Classification rules stored locally in `rules.json`

---

## 📝 Example Workflow

1. **Prepare BOM**: Export from SAP/Oracle as `.xlsx`
2. **Upload**: Drag into Streamlit app
3. **Review**: Check auto-classified items; correct any mistakes
4. **Save**: Click "Save corrections" to train the classifier
5. **Download**: One-click PPTX export
6. **Edit**: Open in PowerPoint for final customization

---

## 🛠️ Development

### Running in Development Mode

```bash
streamlit run app.py --logger.level=debug
```

### Testing with Sample Data

A demo BOM is included:

```bash
# Use in app, or programmatically:
python -c "import json; print(json.load(open('handmade_demo.json')))"
```

### Extending the Classifier

Edit `rules.json` to add new classification patterns:

```json
{
  "PLC": ["S7-1200", "S7-1500", "CPU.*1200", "logo"],
  "HMI": ["KTP.*", "KTP700 Mobile", "PANEL"]
}
```

---

## 📄 License

Proprietary · Emerson Automation Solutions · Pune, India

---

## 📞 Support

**Issues or Questions?**
- Check logs in `./streamlit_logs/` (if enabled)
- Verify all files are in the correct directory
- Ensure Python 3.10+ is installed
- Try resetting `rules.json` if classifications become unstable

---

**Built with ❤️ by Emerson DCS Engineering Team**  
*v3.0 — Image Edition*
