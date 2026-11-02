from excel_utils import unmerge_and_fill

INPUT_FILE = r"C:\Users\nguyeb3\Downloads\CHANGE_BOM_demo.xlsx"


def main():
    try:
        unmerge_and_fill(INPUT_FILE)
    except PermissionError:
        print("Cannot save: close the output file in Excel and run again.")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
