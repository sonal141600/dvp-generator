"""
build_criteria_db.py
────────────────────
Run this ONCE when you add new PDFs to standards_library/
Reads each PDF with pdfplumber (FREE — no tokens for reading)
Uses Claude text (cheap) to extract criteria
Saves to customer_profiles/criteria_db_{customer}.json

Usage:
  python build_criteria_db.py
"""

import anthropic
import json
import os
import re
import pdfplumber

client        = anthropic.Anthropic()
LIBRARY_PATH  = "standards_library"
PROFILES_PATH = "customer_profiles"


def normalize_code(filename: str) -> str:
    """
    Convert filename to standard code.
    TL_1010_EN.pdf     → TL 1010
    VW_50180_EN.pdf    → VW 50180
    DIN_EN_ISO_845.pdf → DIN EN ISO 845
    HES_D_6503.pdf     → HES D 6503
    """
    base = os.path.splitext(filename)[0]
    base = re.sub(r'_EN\b.*$',       '', base)   # remove _EN suffix
    base = re.sub(r'\s*\(\d+\)\s*',  '', base)   # remove (1), (2)
    base = re.sub(r'_\d+$',          '', base)   # trailing version
    base = re.sub(r'_',              ' ', base)   # underscores to spaces
    base = re.sub(r'\s+',            ' ', base)   # normalize spaces
    return base.strip()


def is_standard_filename(filename: str) -> bool:
    """
    Check if filename looks like a standard code vs descriptive name.
    TL_1010_EN.pdf       → True  (standard)
    VW_50180_EN.pdf      → True  (standard)
    PV3900_Components... → False (descriptive — skip)
    """
    base  = os.path.splitext(filename)[0]
    code  = normalize_code(filename)
    words = code.split()
    # If more than 5 words it's probably descriptive not a code
    return len(words) <= 5


def extract_criteria(pdf_path: str,
                     standard_code: str,
                     customer: str) -> str:
    """
    Extract acceptance criteria from standard PDF.
    Step 1: pdfplumber reads text — FREE
    Step 2: Claude extracts criteria — cheap (text only, no vision)
    """
    try:
        # Step 1: Extract text FREE
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            pages = min(len(pdf.pages), 5)
            for i in range(pages):
                page_text = pdf.pages[i].extract_text() or ""
                text += page_text

        if not text.strip():
            print(f"     ⚠️  No text extractable — may be scanned PDF")
            return ""

        # Step 2: Claude extracts criteria (text only — cheap)
        prompt = f"""
You are reading a technical standard document: {standard_code}
Customer: {customer}

From this document extract ONLY the specific acceptance criteria 
or pass/fail values. Keep it SHORT — one sentence maximum.

Examples of good criteria:
- "Burning rate ≤ 100 mm/min"
- "6 cycles without complaint"  
- "Result ≥ 4 after 1 cycle"
- "Gross density 30 ± 3 kg/m³"

Return ONLY the criteria text.
If no clear criteria found return exactly: NONE

DOCUMENT TEXT (first 3000 chars):
{text[:3000]}
"""
        msg = client.messages.create(
            model      = "claude-sonnet-4-20250514",
            max_tokens = 100,
            messages   = [{"role": "user", "content": prompt}]
        )
        result = msg.content[0].text.strip()

        if result.upper() == "NONE" or not result:
            return ""
        return result

    except Exception as e:
        print(f"     ❌ Error: {e}")
        return ""


def build_for_customer(customer: str):
    """
    Build criteria database for one customer.
    Reads all PDFs in standards_library/ and extracts criteria.
    Skips files already in database.
    """
    db_path = os.path.join(
        PROFILES_PATH, f"criteria_db_{customer.lower()}.json"
    )

    # Load existing
    os.makedirs(PROFILES_PATH, exist_ok=True)
    if os.path.exists(db_path):
        with open(db_path, encoding="utf-8") as f:
            db = json.load(f)
        print(f"  Existing entries: {len(db)}")
    else:
        db = {}

    # Scan library
    files = sorted([
        f for f in os.listdir(LIBRARY_PATH)
        if f.lower().endswith(".pdf")
        and not f.startswith(".")
    ])

    print(f"  PDF files found: {len(files)}")
    new   = 0
    skipped = 0

    for filename in files:
        # Skip descriptive filenames
        if not is_standard_filename(filename):
            print(f"  ⏭️  Skip (descriptive): {filename[:50]}")
            skipped += 1
            continue

        code = normalize_code(filename)

        # Skip if already in database
        if code in db:
            skipped += 1
            continue

        print(f"\n  📄 {filename}")
        print(f"     Code: {code}")

        criteria = extract_criteria(
            os.path.join(LIBRARY_PATH, filename),
            code,
            customer
        )

        db[code] = criteria

        if criteria:
            print(f"     ✅ {criteria[:70]}")
            new += 1
        else:
            print(f"     ⚠️  No criteria found — engineer fills manually")

        # Save after every file — don't lose progress
        with open(db_path, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*50}")
    print(f"✅ Done for {customer}!")
    print(f"   New:      {new}")
    print(f"   Skipped:  {skipped}")
    print(f"   Total:    {len(db)}")
    print(f"   Saved to: {db_path}")
    return db


def build_all():
    print("=" * 50)
    print("📚 CRITERIA DATABASE BUILDER")
    print("=" * 50)

    customer = input("\nEnter customer name exactly as it appears "
                     "in the drawing\n(e.g. VOLKSWAGEN, HONDA, "
                     "VINFAST, TATA MOTORS): ").strip()

    if not customer:
        print("❌ No customer name entered")
        return

    db_key  = re.sub(r'[^a-zA-Z0-9]', '_', customer).lower().strip('_')
    db_path = f"customer_profiles/criteria_db_{db_key}.json"

    print(f"\n🏭 Building for: {customer}")
    print(f"   Saving to:    {db_path}")

    build_for_customer(customer)


if __name__ == "__main__":
    build_all()