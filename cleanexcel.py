import os
from openpyxl import load_workbook
from win32com.client import Dispatch
from tqdm import tqdm
import pythoncom


def break_links_with_win32(file_path):
    pythoncom.CoInitialize()  # Necessary for multi-threading
    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(file_path)
    links = wb.LinkSources()  # Get external links

    if links:
        print(f"Total external links to break: {len(links)}")
        for link in tqdm(links, desc="Breaking external links"):
            wb.BreakLink(Name=link, Type=1)  # 1 for external links
    else:
        print("No external links found.")

    wb.Save()
    wb.Close()
    excel.Quit()


def delete_hidden_sheets_openpyxl(file_path):
    wb = load_workbook(file_path)
    hidden_sheets = [
        sheet.title for sheet in wb.worksheets if sheet.sheet_state == "hidden"
    ]

    if not hidden_sheets:
        print("No hidden sheets found.")
        return

    for sheet in hidden_sheets:
        print(f"Deleting hidden sheet: {sheet}")
        del wb[sheet]

    new_file_path = file_path.replace(".xlsx", "_cleaned.xlsx")
    try:
        wb.save(new_file_path)
        print(f"File saved as: {new_file_path}")
    except PermissionError:
        print(
            "Error: Unable to save the file. Please check permissions or if the file is open."
        )
    finally:
        wb.close()


def delete_erroneous_names_with_win32(file_path):
    pythoncom.CoInitialize()
    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(file_path)

    try:
        names = wb.Names
        total_names = names.Count  # Get the total number of names
        print(f"Total names to check: {total_names}")

        for i in tqdm(range(1, total_names + 1), desc="Deleting erroneous names"):
            try:
                name = names.Item(i)
                refers_to = name.RefersTo
                if (
                    "#REF!" in refers_to
                    or "#REF!" in refers_to
                    or "#REF!REF!" in refers_to
                    or "http" in refers_to
                    or "#N/A" in refers_to
                    or "NA()" in refers_to
                    or "#NAME?" in refers_to
                ):
                    name.Delete()
            except Exception as e:
                print(f"Error processing name index {i}: {e}")
    except Exception as e:
        print(f"Error accessing names: {e}")
    finally:
        wb.Save()
        wb.Close()
        excel.Quit()
    print("Erroneous names deleted successfully.")


def main():
    print("Welcome to Enhanced Excel Cleaner")
    file_path = input("Enter the path to the Excel file: ").strip()
    if not os.path.exists(file_path):
        print("File not found. Exiting.")
        return

    print("\nChoose an action:")
    print("1. Break external links (win32com)")
    print("2. Delete hidden sheets (openpyxl)")
    print("3. Delete erroneous names (win32com)")
    print("4. Perform all actions")
    choice = input("\nEnter your choice (1-4): ").strip()

    if choice == "1":
        break_links_with_win32(file_path)
    elif choice == "2":
        delete_hidden_sheets_openpyxl(file_path)
    elif choice == "3":
        delete_erroneous_names_with_win32(file_path)
    elif choice == "4":
        break_links_with_win32(file_path)
        delete_hidden_sheets_openpyxl(file_path)
        delete_erroneous_names_with_win32(file_path)
    else:
        print("Invalid choice. Exiting.")

    print(f"Processing complete for file: {file_path}")


if __name__ == "__main__":
    main()
