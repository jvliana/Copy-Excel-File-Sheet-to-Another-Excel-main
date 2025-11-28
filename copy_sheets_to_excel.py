# Author: Juliana A
# Description:
#     This Python script allows a user to copy all sheets from a source Excel file into a destination Excel file, while preserving all data, formatting, formulas, and layout.
#     It is especially useful when combining multiple workbooks, consolidating reports, or duplicating template sheets with full fidelity.

import xlwings as xw
from tkinter import filedialog, Tk
import os


def get_file_path(prompt_title):
    """Open a file picker and return the selected file path."""
    return filedialog.askopenfilename(
        title=prompt_title,
        filetypes=[("Excel Files", "*.xlsx")]
    )


def get_unique_sheet_name(existing_names, base_name):
    """Generate a unique sheet name if there is a conflict."""
    if base_name not in existing_names:
        return base_name

    counter = 1
    while True:
        new_name = f"{base_name}_copy{counter}"
        if new_name not in existing_names:
            return new_name
        counter += 1


def main():
    # Hide the Tkinter root window
    root = Tk()
    root.withdraw()

    # Show message before file dialogs (ensures instruction appears first)
    print("📂 Select Destination File...", flush=True)
    excel1_path = get_file_path("Select Destination File")
    if not excel1_path:
        print("❌ No Destination File selected. Exiting ❌")
        return

    print("📂 Select Source File...", flush=True)
    excel2_path = get_file_path("Select Source File")
    if not excel2_path:
        print("❌ No Source File selected. Exiting ❌")
        return

    excel1_name = os.path.basename(excel1_path)
    excel2_name = os.path.basename(excel2_path)

    # Open Excel silently
    app = xw.App(visible=False)
    try:
        wb_dest = app.books.open(excel1_path)
        wb_source = app.books.open(excel2_path)

        # Copy each sheet from source to destination
        for sheet in wb_source.sheets:
            new_name = get_unique_sheet_name(
                [s.name for s in wb_dest.sheets], sheet.name
            )
            sheet.api.Copy(Before=wb_dest.sheets[0].api)
            wb_dest.sheets[0].name = new_name
            print(f"✅ Copied sheet '{sheet.name}' as '{new_name}'")

        # Save destination workbook
        wb_dest.save()
        print(
            f"\n🎉 All sheets from '{excel2_name}' were copied into '{excel1_name}' successfully.")

    except Exception as e:
        print(f"⚠️ Error: {e}")
    finally:
        # Always close workbooks and quit Excel properly
        if 'wb_source' in locals():
            wb_source.close()
        if 'wb_dest' in locals():
            wb_dest.close()
        app.quit()


if __name__ == "__main__":
    main()
