"""Rebuild, recalculate, export previews, and generate all packs for M&M, LTM, CLEAN, and EDELWEISS."""
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from work.models_mm import build_mm_workbook
from work.models_ltm import build_ltm_workbook
from work.models_clean import build_clean_workbook
from work.models_edelweiss import build_edelweiss_workbook

from validate_workbook import recalculate_excel_com, scan_formula_errors_and_missing_caches
from modelmaker import Package
from build_package import compute_file_sha

BUILDERS = [
    ("MM", build_mm_workbook, "MM-Financial-Model.xlsx", "MM-Recalculated.xlsx"),
    ("LTM", build_ltm_workbook, "LTM-Financial-Model.xlsx", "LTM-Recalculated.xlsx"),
    ("CLEAN", build_clean_workbook, "CLEAN-Financial-Model.xlsx", "CLEAN-Recalculated.xlsx"),
    ("EDELWEISS", build_edelweiss_workbook, "EDELWEISS-Financial-Model.xlsx", "EDELWEISS-Recalculated.xlsx"),
]

def main():
    print("=== Step 1: Building all 4 workbooks ===")
    for key, builder_fn, orig_name, recalc_name in BUILDERS:
        comp_dir = ROOT / "companies" / key
        comp_dir.mkdir(parents=True, exist_ok=True)
        raw_path = comp_dir / orig_name
        builder_fn(raw_path)
        print(f"Built raw workbook for {key} at {raw_path}")

        cand_dir = ROOT / "candidates" / f"run_native_{key.lower()}"
        cand_dir.mkdir(parents=True, exist_ok=True)
        recalc_cand = cand_dir / recalc_name
        
        # Remove old recalculated workbook if exists to ensure clean rebuild
        if recalc_cand.exists():
            recalc_cand.unlink()

        print(f"Recalculating {key} via Excel COM into {recalc_cand}...")
        recalculate_excel_com(raw_path, recalc_cand)
        sha = compute_file_sha(recalc_cand)
        print(f"Recalculated {key}: SHA-256 = {sha}")

        # Copy to company dir
        comp_recalc = comp_dir / recalc_name
        shutil.copy2(recalc_cand, comp_recalc)

        # Validate with Package
        pkg = Package(recalc_cand.read_bytes())
        errs, missing = scan_formula_errors_and_missing_caches(pkg)
        print(f"Validation for {key}: errors={len(errs)}, missing_caches={len(missing)}")
        if errs or missing:
            raise RuntimeError(f"Validation failed for {key}: {errs}, {missing}")

    print("\n=== Step 2: Exporting Previews and Manifests ===")
    from work.export_previews_and_manifests import export_all
    export_all()

    print("\n=== Step 3: Generating Documentation, Registries, and Packs ===")
    from work.generate_documentation_and_packs import generate_all
    generate_all()

    print("\n=== Rebuild and verification completed successfully! ===")

if __name__ == "__main__":
    main()
