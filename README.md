# 🔍 IP India Trademark Scraper

Automated scrapers for the IP India portals. This project includes two main scripts:
1. **Public Search Scraper** (`scraper.py`) - Searches for trademarks by name.
2. **E-Register Scraper** (`eregister_scraper.py`) - Retrieves full details by Application Number.

## Features

- **Auto CAPTCHA** — Solves CAPTCHAs automatically (no manual intervention)
- **Resume Support** — Safe to stop anytime with `Ctrl+C`, resumes from where it left off via checkpoints
- **Bulk Scraping** — Processes hundreds of trademarks unattended
- **Fast JS Extraction** — Uses JavaScript-based DOM reading for speed
- **CSV Output** — Clean, structured output ready for analysis

## Setup

### Prerequisites

- Python 3.8+
- Google Chrome browser
- ChromeDriver (matching your Chrome version)

### Install

```bash
pip install -r requirements.txt
```

## Usage

### 1. Public Search Scraper (`scraper.py`)
Reads trademark names from an Excel file and extracts all matching records.

1. **Prepare your Excel file** — Create `input.xlsx` in the project folder with these columns:
   - `Trademark Name` — the trademark to search
   - `International classes` — class number(s), comma-separated

2. **Run the scraper:**
   ```bash
   python scraper.py
   ```

3. **Output** — Results are saved to `output.csv`. Progress tracked in `checkpoint.txt`.

### 2. E-Register Scraper (`eregister_scraper.py`)
Reads application numbers from an Excel file and extracts detailed trademark data from the E-Register portal.

1. **Prepare your Excel file** — Use `PCPB_Input.xlsx` in the project folder with these columns:
   - `Country`
   - `Application Number`

2. **Run the scraper:**
   ```bash
   python eregister_scraper.py
   ```
   *Note: On the first run in a session, you will need to manually log in with OTP on the IP India site. The script handles the rest.*

3. **Output** — Results are saved to `E-Register_output.csv`. Progress tracked in `eregister_checkpoint.txt`.

## Output Fields

### Public Search Output
| Field | Description |
|---|---|
| Search Keyword | Original search term |
| Search Class | Class searched |
| Application Number | TMR application number |
| Word Mark | Registered word mark |
| Class | Trademark class |
| Status | Registration status |
| Appl. Date | Application date |
| Proprietor | Trademark owner |
| Journal No / Date | Publication journal info |
| Used Since / Valid Upto | Usage and validity dates |
| Goods & Services | Description of goods/services |
| Address | Proprietor address |
| Agent / Attorney | Legal representative |

### E-Register Output
Extensive detailed fields including Application Number, Word Mark, Proprietor Details, Agent Details, Dates, Status, TM Image, and more.

## Resume & Checkpoint

The scrapers track progress in their respective checkpoint files (`checkpoint.txt` and `eregister_checkpoint.txt`). If interrupted:
- **Resume:** Just run the script again. It will skip already processed entries.
- **Restart from scratch:** Delete the checkpoint file and the output CSV.

## License

MIT
