from tkinter import filedialog, Tk, Button, Label, Toplevel
import pandas as pd

def select_file(root):
    root.withdraw()
    filename = filedialog.askopenfilename(title="Select the input Excel file", filetypes=[("Excel files", "*.xlsx")])
    return filename

def select_sheet_from_file(filename, root):
    xl = pd.ExcelFile(filename)
    sheet_names = xl.sheet_names
    return show_sheet_dialog(sheet_names, root)

def show_sheet_dialog(sheet_names, root):
    sheet_selected = [None]

    def on_button_click(sheet):
        sheet_selected[0] = sheet
        top.destroy()
        root.destroy()

    top = Toplevel(root)
    top.title("Select Sheet")
    Label(top, text="Choose a sheet:").pack(pady=10)

    for sheet in sheet_names:
        Button(top, text=sheet, command=lambda s=sheet: on_button_click(s)).pack(pady=5)

    top.mainloop()

    return sheet_selected[0]

    """
    create a function that gives the option to choose which projects 
    from the lables column to create new sheets with counts for
    """