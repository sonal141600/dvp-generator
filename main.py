# main.py
# ─────────────────────────────────────────────────────────────
# Processes all RFQ folders and generates DVP Test Plans.
# Usage:
#   python main.py                  → process all folders in rfq_inputs/
#   python main.py rfq_inputs/RFQ_Toyota_002  → process one folder
# ─────────────────────────────────────────────────────────────

import os
import sys
import time
from dvp_reader import generate_dvp

STANDARDS_LIBRARY = "standards_library"
RFQ_BASE          = "rfq_inputs"
OUTPUT_BASE       = "output"


def process_rfq(folder_path: str):
    part_no     = os.path.basename(folder_path).replace(" ", "_")
    output_path = os.path.join(OUTPUT_BASE, f"DVP_{part_no}.xlsx")

    print(f"\n{'='*60}")
    print(f"📂 Processing: {folder_path}")
    print(f"{'='*60}")

    start = time.time()
    try:
        tests = generate_dvp(
            folder_path       = folder_path,
            standards_library = STANDARDS_LIBRARY,
            output_path       = output_path,
        )
        elapsed = round(time.time() - start, 1)
        if tests:
            available   = sum(1 for t in tests if t["available"])
            unavailable = len(tests) - available
            return {
                "folder":      folder_path,
                "output":      output_path,
                "total":       len(tests),
                "available":   available,
                "unavailable": unavailable,
                "time":        elapsed,
                "status":      "ok",
            }
        else:
            return {"folder": folder_path, "status": "empty", "time": elapsed}

    except Exception as e:
        elapsed = round(time.time() - start, 1)
        print(f"❌ Error: {e}")
        return {"folder": folder_path, "status": "error", "error": str(e), "time": elapsed}


def print_summary(results: list):
    print(f"\n{'='*60}")
    print("📊 BATCH SUMMARY")
    print(f"{'='*60}")

    ok      = [r for r in results if r["status"] == "ok"]
    empty   = [r for r in results if r["status"] == "empty"]
    errors  = [r for r in results if r["status"] == "error"]

    for r in ok:
        print(f"  ✅ {os.path.basename(r['folder']):<30} "
              f"{r['total']} tests  "
              f"🟢{r['available']} 🔴{r['unavailable']}  "
              f"({r['time']}s)")

    for r in empty:
        print(f"  ⚠️  {os.path.basename(r['folder']):<30} No tests found ({r['time']}s)")

    for r in errors:
        print(f"  ❌ {os.path.basename(r['folder']):<30} ERROR: {r.get('error', '')} ({r['time']}s)")

    print(f"\n  Total: {len(results)} folders | "
          f"✅ {len(ok)} ok | "
          f"⚠️  {len(empty)} empty | "
          f"❌ {len(errors)} errors")


if __name__ == "__main__":

    # Single folder mode
    if len(sys.argv) > 1:
        folder = sys.argv[1]
        if not os.path.isdir(folder):
            print(f"❌ Not a folder: {folder}")
            sys.exit(1)
        result = process_rfq(folder)
        print_summary([result])

    # Batch mode — all folders in rfq_inputs/
    else:
        if not os.path.isdir(RFQ_BASE):
            print(f"❌ RFQ base folder not found: {RFQ_BASE}")
            sys.exit(1)

        folders = sorted([
            os.path.join(RFQ_BASE, f)
            for f in os.listdir(RFQ_BASE)
            if os.path.isdir(os.path.join(RFQ_BASE, f))
            and not f.startswith(".")
        ])

        if not folders:
            print(f"⚠️  No folders found in {RFQ_BASE}/")
            sys.exit(0)

        print(f"🔍 Found {len(folders)} RFQ folder(s) in '{RFQ_BASE}/'")
        os.makedirs(OUTPUT_BASE, exist_ok=True)

        results = []
        for folder in folders:
            result = process_rfq(folder)
            results.append(result)

        print_summary(results)