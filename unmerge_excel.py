from workbook import unmerge_and_fill

from config import UNMERGE_FILE as INPUT_FILE


def main():
    try:
        unmerge_and_fill(INPUT_FILE)
    except PermissionError:
        print("Cannot save: close the output file in Excel and run again.")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
