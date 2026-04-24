import anthropic
import base64
import json
import os
import re
from PIL import Image
import io

Image.MAX_IMAGE_PIXELS = None
client = anthropic.Anthropic()


# ── Prompts ────────────────────────────────────────────────────────────────────

SCOUT_PROMPT = """
Look at this strip of an engineering drawing.
Reply ONLY with valid JSON — no markdown, no explanation:
{
  "has_bom_table":      true or false,
  "has_title_block":    true or false,
  "has_dimensions":     true or false,
  "has_material_specs": true or false,
  "description":        "one short sentence"
}

has_bom_table      = true if you see a table with rows/columns 
                     containing part names, materials, weights.
has_title_block    = true if you see a structured block with 
                     part number, part name, revision, scale.
has_dimensions     = true if you see large overall dimension 
                     lines (values > 500mm) spanning the full part.
has_material_specs = true if you see GSM / g/m² / material 
                     specification text.
"""

EXTRACT_PROMPT = """
You are an expert automotive engineering drawing reader.
Extract data from this drawing strip.

IMPORTANT RULES:
- Output ENGLISH ONLY. If text is in German AND English, use English.
  If text is German only, translate it to English.
- Do NOT output any German words, unicode escapes like \\u00fc, 
  or mixed language text.
- Copy numbers exactly as shown.

════════════════════════════════════════
WHAT TO EXTRACT:
════════════════════════════════════════

1. TITLE BLOCK (bottom-right corner):
   - part_number:  e.g. "2FK.863.021"  (Teil-Nr / Part no field)
   - part_name:    English name only e.g. "FLOOR MAT"
                   (Benennung / Designation field)
   - customer:     company name e.g. "VOLKSWAGEN"
   - revision:     letter e.g. "A"
   - date:         drawing date
   - scale:        e.g. "1:2"

2. OVERALL DIMENSIONS:
   - Look for large dimension lines spanning the FULL part
   - Values will be large numbers like 1813, 900, 300
   - They may appear in brackets: (1813)
   - overall_L_mm: longest dimension (length)
   - overall_W_mm: width dimension
   - overall_H_mm: height or thickness

3. PARTS / MATERIALS TABLE:
   Table has these columns (may be German/English):
   Field | Item | Part-No | Title/Name | 
   Semi-finished Material | Material Treatment | 
   Surface Treatment | Qty | Weight(g) | Material Marking

   For EACH ROW extract:
   - item_no:    e.g. "1A", "1B", "2A", "3A"
   - part_number: Teil-Nr value, write "o.Z." as "w/o drawing"
   - part_name:  English name only
                 e.g. "Sound absorber floor cover RL"
                 NOT "Daempfung Bodenbelag-RL"
   - qty:        quantity number
   - weight_g:   weight in grams (integer from Gewicht column)
   - material:   English material description
                 e.g. "PUR foam acc. to TL 52602, 
                       density 55±5.5 kg/m3, color gray"
                 NOT German text
   - gsm:        surface weight in g/m²
                 Look for "Flaechengewicht" or "surface weight"
                 Take the TOTAL/gesamt value
                 e.g. "total surface weight: 1310 g/m²" → 1310
                 If layers listed, sum them or take gesamt value
   - remarks:    material marking e.g. ">PUR<", ">PET<", ">PP<"
   CRITICAL FOR WEIGHTS:
- "Gewicht / Weight (g)" column has integer values like 2456, 64
- ALWAYS extract weight_g for every row that has a number there
- Never leave weight_g as null if a number is visible in that column

CRITICAL FOR GSM:
- GSM only applies to textile/fleece layers (item 2A, 2B etc.)
- For PUR foam items — gsm should be null
- Look for "Flaechengewicht gesamt" or "total surface weight" value
- e.g. "Halbzeug gesamt Flaechengewicht: (1310±130) g/m²" → gsm: 1310

════════════════════════════════════════
Return ONLY valid JSON, no other text:
════════════════════════════════════════
{
    "customer": "",
    "drawing_number": "",
    "part_number": "",
    "part_name": "",
    "revision": "",
    "date": "",
    "scale": "",
    "project_model": "",
    "weight_g": null,
    "raw_material": "",
    "gsm": null,
    "color_code": "",
    "overall_L_mm": null,
    "overall_W_mm": null,
    "overall_H_mm": null,
    "parts_table": [
        {
            "item_no": "",
            "part_number": "",
            "part_name": "",
            "qty": null,
            "material": "",
            "gsm": null,
            "weight_g": null,
            "remarks": ""
        }
    ],
    "notes": [],
    "confidence": "high or medium or low"
}
"""


# ── Image helpers ──────────────────────────────────────────────────────────────

def _img_to_b64(img: Image.Image, quality: int = 90) -> str:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=quality)
    return base64.standard_b64encode(buf.getvalue()).decode()


def _ask_claude(b64: str, prompt: str, max_tokens: int = 300) -> dict:
    """Send image to Claude and return parsed JSON."""
    msg = client.messages.create(
        model      = "claude-sonnet-4-20250514",
        max_tokens = max_tokens,
        messages   = [{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type":       "base64",
                        "media_type": "image/jpeg",
                        "data":       b64,
                    }
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'\s*```$',     '', raw)
    return json.loads(raw)


# ── TIF reader ─────────────────────────────────────────────────────────────────

def _read_tif(file_path: str) -> dict:
    print(f"\n  📄 TIF: {os.path.basename(file_path)}")

    img = Image.open(file_path)
    W, H = img.size
    # Crop to bottom 55% only — all data lives there
    # Top half is just engineering views with no BOM data
    img = img.crop((0, int(H * 0.35), W, H))
    W, H = img.size
    print(f"  ✂️  Cropped to bottom 55%: {W}x{H}px")
    
    print(f"  📐 Size: {W}x{H}px")

    # Determine strips based on aspect ratio
    aspect = W / H
    if aspect > 4:
        n = 5      # very wide — VW/Skoda
    elif aspect > 1.5:
        n = 4      # landscape
    else:
        n = 3      # portrait

    print(f"  📊 Aspect {aspect:.1f} → {n} strips")

    strips = []
    overlap = int(W * 0.10)  # 10% overlap on each side
    for i in range(n):
        x1 = max(0, int(W * i / n) - overlap)
        x2 = min(W, int(W * (i+1) / n) + overlap)
        crop = img.crop((x1, 0, x2, H))
        crop.thumbnail((3000, 3000), Image.LANCZOS)
        strips.append({"index": i+1, "image": crop})

    # Pass 1: Scout
    print(f"\n  🔎 Pass 1 — scouting {n} strips...")
    for s in strips:
        b64 = _img_to_b64(s["image"], quality=70)
        try:
            s["scout"] = _ask_claude(b64, SCOUT_PROMPT, max_tokens=200)
        except Exception as e:
            print(f"    ⚠️  Scout error strip {s['index']}: {e}")
            s["scout"] = {}

        flags = []
        if s["scout"].get("has_bom_table"):      flags.append("BOM")
        if s["scout"].get("has_title_block"):    flags.append("Title")
        if s["scout"].get("has_dimensions"):     flags.append("Dims")
        if s["scout"].get("has_material_specs"): flags.append("Mats")

        icon = "✅" if flags else "⬜"
        desc = s["scout"].get("description", "")[:55]
        print(f"    {icon} Strip {s['index']}: "
              f"{', '.join(flags) or 'views only'} — {desc}")

    # Select relevant strips
    relevant = [
        s for s in strips
        if (s["scout"].get("has_bom_table")      or
            s["scout"].get("has_title_block")    or
            s["scout"].get("has_material_specs") or
            s["scout"].get("has_dimensions"))
    ]

    if not relevant:
        print("  ⚠️  Nothing flagged — extracting all strips")
        relevant = strips

    # Rank by number of flags — take only top 2 strips
    def _read_tif(file_path: str) -> dict:
      print(f"\n  📄 TIF: {os.path.basename(file_path)}")

    img = Image.open(file_path)
    W, H = img.size
    print(f"  📐 Size: {W}x{H}px")

    # Determine strips based on aspect ratio
    aspect = W / H
    if aspect > 4:
        n = 5
    elif aspect > 1.5:
        n = 4
    else:
        n = 3

    print(f"  📊 Aspect {aspect:.1f} → {n} strips")

    strips = []
    for i in range(n):
        x1   = int(W * i / n)
        x2   = int(W * (i+1) / n)
        crop = img.crop((x1, 0, x2, H))
        crop.thumbnail((3000, 3000), Image.LANCZOS)
        strips.append({"index": i+1, "image": crop})

    # Pass 1: Scout
    print(f"\n  🔎 Pass 1 — scouting {n} strips...")
    for s in strips:
        b64 = _img_to_b64(s["image"], quality=70)
        try:
            s["scout"] = _ask_claude(b64, SCOUT_PROMPT, max_tokens=200)
        except Exception as e:
            print(f"    ⚠️  Scout error strip {s['index']}: {e}")
            s["scout"] = {}

        flags = []
        if s["scout"].get("has_bom_table"):      flags.append("BOM")
        if s["scout"].get("has_title_block"):    flags.append("Title")
        if s["scout"].get("has_dimensions"):     flags.append("Dims")
        if s["scout"].get("has_material_specs"): flags.append("Mats")

        icon = "✅" if flags else "⬜"
        desc = s["scout"].get("description", "")[:55]
        print(f"    {icon} Strip {s['index']}: "
              f"{', '.join(flags) or 'views only'} — {desc}")

    # Select relevant strips
    relevant = [
        s for s in strips
        if (s["scout"].get("has_bom_table")      or
            s["scout"].get("has_title_block")    or
            s["scout"].get("has_material_specs") or
            s["scout"].get("has_dimensions"))
    ]

    if not relevant:
        print("  ⚠️  Nothing flagged — extracting all strips")
        relevant = strips

    # Rank strips — BOM+Mats together scores highest
    def flag_count(s):
        sc   = s["scout"]
        bom  = bool(sc.get("has_bom_table"))
        mats = bool(sc.get("has_material_specs"))
        titl = bool(sc.get("has_title_block"))
        dims = bool(sc.get("has_dimensions"))
        score = (bom and mats) * 3 + titl + dims + bom + mats
        return score

    relevant = sorted(relevant, key=flag_count, reverse=True)

    # Always include last strip (title block is always rightmost)
    last_strip = strips[-1]
    top = relevant[:1]
    if last_strip not in top:
        top = top + [last_strip]
    relevant = top

    print(f"  🎯 Top strips selected: {[s['index'] for s in relevant]}")

    # Pass 2: Extract
    print(f"\n  🔬 Pass 2 — extracting {len(relevant)}/{n} strips...")
    results = []
    for s in relevant:
        desc = s["scout"].get("description", "")[:50]
        print(f"    🔬 Strip {s['index']}: {desc}...")

        b64 = _img_to_b64(s["image"], quality=95)
        try:
            result = _ask_claude(b64, EXTRACT_PROMPT, max_tokens=2500)
            pn   = result.get("part_number", "")
            gsm  = result.get("gsm")
            pts  = len(result.get("parts_table", []))
            conf = result.get("confidence", "")
            print(f"       Part#: {pn} | GSM: {gsm} | "
                  f"Parts: {pts} | Conf: {conf}")
            results.append(result)

        except json.JSONDecodeError:
            print(f"    🔄 JSON error — retrying strip {s['index']}...")
            try:
                retry = (EXTRACT_PROMPT +
                         "\n\nCRITICAL: Return ONLY valid JSON. "
                         "No trailing commas. No comments.")
                result = _ask_claude(b64, retry, max_tokens=2500)
                results.append(result)
                print(f"    ✅ Retry succeeded")
            except Exception as e2:
                print(f"    ❌ Retry failed: {e2}")

        except Exception as e:
            print(f"    ❌ Error: {e}")

    return _merge(results)

# ── PDF reader ─────────────────────────────────────────────────────────────────

def _read_pdf(file_path: str) -> dict:
    print(f"\n  📄 PDF: {os.path.basename(file_path)}")

    try:
        import pdfplumber
        results = []

        with pdfplumber.open(file_path) as pdf:
            pages = min(len(pdf.pages), 5)
            for i in range(pages):
                print(f"    📃 Page {i+1}...")
                text = pdf.pages[i].extract_text() or ""

                if len(text) > 100:
                    result = _extract_text(text, i+1)
                else:
                    img = pdf.pages[i].to_image(resolution=150).original
                    b64 = _img_to_b64(img)
                    result = _ask_claude(b64, EXTRACT_PROMPT, max_tokens=2500)

                if result:
                    print(f"       {result.get('part_name','?')} | "
                          f"conf: {result.get('confidence','?')}")
                    results.append(result)

        return _merge(results)

    except Exception as e:
        print(f"  ❌ PDF error: {e}")
        return _empty()


def _extract_text(text: str, page: int) -> dict:
    prompt = f"""
Extract BOM data from this engineering document.
Output ENGLISH ONLY — translate German text if needed.
No unicode escapes like \\u00fc.

TEXT:
{text}

Return ONLY valid JSON:
{{
    "customer": "",
    "drawing_number": "",
    "part_number": "",
    "part_name": "",
    "revision": "",
    "date": "",
    "scale": "",
    "project_model": "",
    "weight_g": null,
    "raw_material": "",
    "gsm": null,
    "color_code": "",
    "overall_L_mm": null,
    "overall_W_mm": null,
    "overall_H_mm": null,
    "parts_table": [],
    "notes": [],
    "confidence": "high or medium or low"
}}
"""
    try:
        msg = client.messages.create(
            model      = "claude-sonnet-4-20250514",
            max_tokens = 2000,
            messages   = [{"role": "user", "content": prompt}]
        )
        raw = msg.content[0].text.strip()
        raw = re.sub(r'^```json\s*', '', raw)
        raw = re.sub(r'\s*```$',     '', raw)
        return json.loads(raw)
    except Exception as e:
        print(f"    ❌ Text extract error p{page}: {e}")
        return {}


# ── Merge ──────────────────────────────────────────────────────────────────────

def _merge(results: list) -> dict:
    if not results:
        return _empty()
    if len(results) == 1:
        return results[0]

    base = results[0].copy()

    for r in results[1:]:
        if not r:
            continue
        for key, val in r.items():

            if key == "parts_table":
                existing = base.get("parts_table") or []
                seen = {p.get("item_no") for p in existing}
                for p in (val or []):
                    k = p.get("item_no")
                    if k and k not in seen:
                        existing.append(p)
                        seen.add(k)
                base["parts_table"] = existing

            elif key == "notes":
                base["notes"] = list(set(
                    (base.get("notes") or []) + (val or [])
                ))

            elif key == "confidence":
                rank = {"high": 0, "medium": 1, "low": 2}
                cur  = rank.get(base.get("confidence", "low"), 2)
                nw   = rank.get(val or "low", 2)
                base["confidence"] = ["high","medium","low"][max(cur, nw)]

            else:
                if base.get(key) in (None, "", [], {}):
                    if val not in (None, "", [], {}):
                        base[key] = val

    return base


def _empty() -> dict:
    return {
        "customer":      "",
        "drawing_number":"",
        "part_number":   "",
        "part_name":     "",
        "revision":      "",
        "date":          "",
        "scale":         "",
        "project_model": "",
        "weight_g":      None,
        "raw_material":  "",
        "gsm":           None,
        "color_code":    "",
        "overall_L_mm":  None,
        "overall_W_mm":  None,
        "overall_H_mm":  None,
        "parts_table":   [],
        "notes":         [],
        "confidence":    "low",
    }


# ── Public API ─────────────────────────────────────────────────────────────────

def read_drawing(file_path: str) -> dict:
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".tif", ".tiff", ".png", ".jpg", ".jpeg"):
        return _read_tif(file_path)
    elif ext == ".pdf":
        return _read_pdf(file_path)
    else:
        print(f"  ⚠️  Unsupported: {ext}")
        return _empty()


def read_folder(folder_path: str) -> list:
    """
    Read all TIF and PDF files in folder.
    Each TIF = one part = one result.
    """
    print(f"\n📂 Folder: {folder_path}")

    tifs = sorted([
        f for f in os.listdir(folder_path)
        if f.lower().endswith((".tif", ".tiff"))
        and not f.startswith(".")
    ])
    pdfs = sorted([
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
        and not f.startswith(".")
    ])

    print(f"  TIFs: {tifs}")
    print(f"  PDFs: {pdfs}")

    results = []

    for f in tifs:
        result = read_drawing(os.path.join(folder_path, f))
        result["source_file"] = f
        results.append(result)

    # Skip PDFs for now — TDO spec docs have no BOM data
    # Uncomment below when PDF reading is needed
    # for f in pdfs:
    #     result = read_drawing(os.path.join(folder_path, f))
    #     result["source_file"] = f
    #     results.append(result)
    if pdfs:
        print(f"  ⏭️  Skipping PDFs: {pdfs}")

    print(f"\n✅ {len(results)} file(s) processed")
    return results


# ── Test ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":

    folder = "rfq_inputs/RFQ_VW_001"

    if not os.path.exists(folder):
        print(f"❌ Not found: {folder}")
        raise SystemExit(1)

    print("🚀 drawing_reader.py")
    print("=" * 60)

    results = read_folder(folder)

    print("\n" + "=" * 60)
    print("📊 RESULTS")
    print("=" * 60)

    for i, r in enumerate(results):
        src = r.get("source_file", f"file {i+1}")
        print(f"\n{'─'*60}")
        print(f"File:        {src}")
        print(f"Customer:    {r.get('customer')}")
        print(f"Part No:     {r.get('part_number')}")
        print(f"Part Name:   {r.get('part_name')}")
        print(f"Revision:    {r.get('revision')}")
        print(f"Date:        {r.get('date')}")
        print(f"Scale:       {r.get('scale')}")
        print(f"Material:    {r.get('raw_material')}")
        print(f"GSM:         {r.get('gsm')}")
        print(f"Weight:      {r.get('weight_g')}g")
        print(f"Dimensions:  {r.get('overall_L_mm')} x "
              f"{r.get('overall_W_mm')} x "
              f"{r.get('overall_H_mm')} mm")
        print(f"Confidence:  {r.get('confidence')}")

        parts = r.get("parts_table", [])
        if parts:
            print(f"\n  Parts ({len(parts)} items):")
            print(f"  {'Item':<5} {'Part Name':<30} "
                  f"{'Material':<35} {'GSM':>6} {'Wt(g)':>7} {'Mark'}")
            print(f"  {'─'*5} {'─'*30} {'─'*35} "
                  f"{'─'*6} {'─'*7} {'─'*8}")
            for p in parts:
                print(
                    f"  {str(p.get('item_no','')):<5} "
                    f"{str(p.get('part_name',''))[:30]:<30} "
                    f"{str(p.get('material',''))[:35]:<35} "
                    f"{str(p.get('gsm') or '')[:6]:>6} "
                    f"{str(p.get('weight_g') or '')[:7]:>7} "
                    f"{str(p.get('remarks',''))[:8]}"
                )
        else:
            print("  No parts table found")

        notes = r.get("notes", [])
        if notes:
            print(f"\n  Notes: {len(notes)} found")

    # Save
    os.makedirs("output", exist_ok=True)
    out = "output/drawing_extraction_raw.json"
    with open(out, "w") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\n💾 Saved: {out}")