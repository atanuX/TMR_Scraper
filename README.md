# 🔍 IP India Trademark Scraper

Automated scraper for the [IP India Public Trademark Search](https://tmrsearch.ipindia.gov.in/tmrpublicsearch/) portal. Reads trademark names from an Excel file and extracts all matching records with full details.

## Features

- **Auto CAPTCHA** — Solves CAPTCHAs automatically (no manual intervention)
- **Resume Support** — Safe to stop anytime with `Ctrl+C`, resumes from where it left off
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

1. **Prepare your Excel file** — Create `input.xlsx` in the project folder with these columns:
   - `Trademark Name` — the trademark to search
   - `International classes` — class number(s), comma-separated

2. **Run the scraper:**
   ```bash
   python scraper.py
   ```

3. **Output** — Results are saved to `output.csv`

### Example Excel Format

| Trademark Name | International classes |
|---|---|
| SUPERIA | 3 |
| FIAMA | 3 |
| VIVEL | 3,5 |

## Output Fields

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

## Resume & Checkpoint

The scraper tracks progress in `checkpoint.txt`. If interrupted:
- **Resume:** Just run `python scraper.py` again
- **Restart from scratch:** Delete `checkpoint.txt` and `output.csv`

## License

MIT
