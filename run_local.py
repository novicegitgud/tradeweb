import argparse
from contextlib import ExitStack
from datetime import date
from pathlib import Path

from Tradeweb import (
    DEFAULT_ADDITIONAL_DATA_PATH,
    DEFAULT_TEMPLATE_PATH,
    build_zip,
    generate_outputs,
    next_weekday,
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Run the Tradeweb generator directly against an input folder."
    )
    parser.add_argument(
        "--input-dir",
        default="input",
        help="Folder containing PRINOS/PORTFELJ XML files and optional Excel overrides.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Folder where generated CSV files, ZIP, and log will be written.",
    )
    parser.add_argument(
        "--portfolio-date",
        default=date.today().isoformat(),
        help="Portfolio date in YYYY-MM-DD format. Defaults to today.",
    )
    parser.add_argument(
        "--trade-date",
        help="Trade date in YYYY-MM-DD format. Defaults to the next weekday after portfolio date.",
    )
    parser.add_argument(
        "--template",
        help="Optional path to a template.xlsx override.",
    )
    parser.add_argument(
        "--additional",
        help="Optional path to an ADDITIONAL_DATA.xlsx override.",
    )
    return parser.parse_args()


def choose_existing_path(*candidates: Path) -> Path:
    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate
    raise FileNotFoundError(f"None of these files exist: {', '.join(str(c) for c in candidates if c)}")


def discover_inputs(input_dir: Path) -> tuple[Path, list[Path]]:
    xml_files = sorted(input_dir.glob("*.xml"))
    prinos_files = [path for path in xml_files if "PRINOS" in path.name.upper()]
    portfelj_files = [path for path in xml_files if "PORTFELJ" in path.name.upper()]

    if len(prinos_files) != 1:
        raise ValueError(
            f"Expected exactly one PRINOS XML file in {input_dir}, found {len(prinos_files)}."
        )

    if not portfelj_files:
        raise ValueError(f"No PORTFELJ XML files found in {input_dir}.")

    return prinos_files[0], portfelj_files


def main():
    args = parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    portfolio_date = date.fromisoformat(args.portfolio_date)
    trade_date = date.fromisoformat(args.trade_date) if args.trade_date else next_weekday(portfolio_date)

    template_path = choose_existing_path(
        Path(args.template).resolve() if args.template else None,
        input_dir / "template.xlsx",
        DEFAULT_TEMPLATE_PATH,
    )
    additional_path = choose_existing_path(
        Path(args.additional).resolve() if args.additional else None,
        input_dir / "ADDITIONAL_DATA.xlsx",
        DEFAULT_ADDITIONAL_DATA_PATH,
    )
    prinos_path, portfelj_paths = discover_inputs(input_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    with ExitStack() as stack:
        prinos_file = stack.enter_context(prinos_path.open("rb"))
        portfelj_files = [stack.enter_context(path.open("rb")) for path in portfelj_paths]

        outputs, log_df = generate_outputs(
            template_source=template_path,
            additional_source=additional_path,
            prinos_source=prinos_file,
            portfelj_files=portfelj_files,
            portfolio_date=portfolio_date,
            trade_date=trade_date,
        )

    for filename, content in outputs.items():
        (output_dir / filename).write_bytes(content)

    zip_path = output_dir / f"tradeweb_csvs_{trade_date.strftime('%Y%m%d')}.zip"
    zip_path.write_bytes(build_zip(outputs))

    log_path = output_dir / "processing_log.csv"
    log_df.to_csv(log_path, index=False, encoding="utf-8-sig")

    print(f"Created {len(outputs)} CSV file(s).")
    print(f"Processing log: {log_path}")
    print(f"ZIP archive: {zip_path}")


if __name__ == "__main__":
    main()
