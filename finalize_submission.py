from append_baseline_metrics import (
    BASELINE_SHEET,
    WORKBOOK_PATH,
    append_baseline_metrics,
)


def main() -> None:
    try:
        append_baseline_metrics(WORKBOOK_PATH)
    except PermissionError:
        print(
            f"Could not update {WORKBOOK_PATH.name} because it is currently in use. "
            "Close the workbook in Microsoft Excel, then run this script again."
        )
        return

    print(f"Appended styled '{BASELINE_SHEET}' sheet to {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()
