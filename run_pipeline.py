from __future__ import annotations

import argparse
from pathlib import Path
import sys

import pandas as pd

script_dir_candidates = []
if "__file__" in globals():
    script_dir_candidates.append(Path(__file__).resolve().parent)
script_dir_candidates.extend(
    [
        Path("/content/Bridgestone-FCL"),
        Path.cwd(),
    ]
)
for candidate in script_dir_candidates:
    if candidate.exists() and str(candidate) not in sys.path:
        sys.path.insert(0, str(candidate))

from convert_extracted_to_layout import write_layout_xlsx
from extract_to_df import (
    DEFAULT_SHEET,
    ask_user_for_files,
    ask_user_for_sheet,
    clean_rate_card_df,
    parse_rate_card_sheet,
)

DEFAULT_CODE_ROOT = Path("/content/Bridgestone-FCL")
DEFAULT_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team/"
    "Documents/AI Adoption RMT/RMT_Bridgestone/FCL"
)
DEFAULT_INPUT_DIR = DEFAULT_DATA_ROOT / "input"
DEFAULT_PROCESSING_DIR = DEFAULT_DATA_ROOT / "processing"
DEFAULT_OUTPUT_DIR = DEFAULT_DATA_ROOT / "output"


def load_single_input_file(file_path: Path) -> pd.DataFrame:
    """Load a single input workbook into a cleaned extracted dataframe."""
    xls = pd.ExcelFile(file_path)
    if DEFAULT_SHEET in xls.sheet_names:
        sheet_to_read = DEFAULT_SHEET
    else:
        sheet_to_read = ask_user_for_sheet(file_path, xls.sheet_names)
        if sheet_to_read is None:
            raise ValueError(f"No sheet selected for {file_path.name}")

    df, currency = parse_rate_card_sheet(file_path, sheet_to_read)
    df = clean_rate_card_df(df)
    if currency:
        df["currency"] = currency
    return df


def run_pipeline(input_dir: Path, processing_dir: Path, output_dir: Path) -> None:
    """Run end-to-end processing: input -> extracted -> layout output."""
    input_files = sorted(input_dir.glob("*.xlsx"))
    if not input_files:
        raise FileNotFoundError(f"No .xlsx files found in {input_dir}")

    selected_files = ask_user_for_files(input_files)
    processing_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    for source_file in selected_files:
        print(f"\nProcessing: {source_file.name}")
        extracted_df = load_single_input_file(source_file)

        extracted_output = processing_dir / f"{source_file.stem}_extracted.xlsx"
        extracted_df.to_excel(extracted_output, index=False)

        layout_output = output_dir / f"{source_file.stem}_extracted_layout.xlsx"
        write_layout_xlsx(extracted_df, layout_output)

        print(f"Extracted saved to: {extracted_output}")
        print(f"Layout saved to: {layout_output}")

    print("\nPipeline completed.")


def parse_args() -> argparse.Namespace:
    """Parse CLI args for local or Colab/Drive usage."""
    parser = argparse.ArgumentParser(description="Run FCL end-to-end pipeline.")
    parser.add_argument(
        "--base-dir",
        type=Path,
        default=DEFAULT_CODE_ROOT,
        help="Base directory containing input/processing/output folders.",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=None,
        help="Override input directory.",
    )
    parser.add_argument(
        "--processing-dir",
        type=Path,
        default=None,
        help="Override processing directory.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Override output directory.",
    )
    args, _unknown = parser.parse_known_args()
    return args


if __name__ == "__main__":
    args = parse_args()
    input_dir = args.input_dir or DEFAULT_INPUT_DIR
    processing_dir = args.processing_dir or DEFAULT_PROCESSING_DIR
    output_dir = args.output_dir or DEFAULT_OUTPUT_DIR
    run_pipeline(input_dir=input_dir, processing_dir=processing_dir, output_dir=output_dir)
