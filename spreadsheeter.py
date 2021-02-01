from openpyxl import Workbook, load_workbook
from openpyxl.utils import FORMULAE
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename

EXECLS_PATHS = {"Table": "", "Prices": ""}

OUT_WORKBOOK_SUFFIX = "_out"
DEBUG_LABEL = None


def update_debug_label(txt, color="red"):
    DEBUG_LABEL.config(fg=color, text=txt)


def get_out_filepath():
    prefix, delimiter, postfix = EXECLS_PATHS["Table"].rpartition(".")
    return prefix + "_out" + delimiter + postfix


def get_price_of_food_id(prices_sheet, id):
    for row in prices_sheet.iter_rows(min_row=2, max_row=prices_sheet.max_row, max_col=prices_sheet.max_column):
        if(row[0].value == id):
            return row[-1].value

    update_debug_label("There is anid in the table that is not in the prices!")


def calculate_spreadsheet():
    if(EXECLS_PATHS["Table"] == "" or EXECLS_PATHS["Prices"] == ""):
        update_debug_label("You didn't select a table!")
        return

    table_workbook = load_workbook(filename=EXECLS_PATHS["Table"])
    table_sheet = table_workbook.active

    prices_workbook = load_workbook(filename=EXECLS_PATHS["Prices"])
    prices_sheet = prices_workbook.active

    for col in table_sheet.iter_cols(min_row=1, min_col=2, max_row=table_sheet.max_row, max_col=table_sheet.max_column):
        col_food_price = get_price_of_food_id(prices_sheet, col[0].value)
        for cell in col:
            if cell.row == 1 or cell.row == 2:
                continue
            cell_value = cell.value if cell.value else 0
            cell.value = float(cell_value)*float(col_food_price)

    for row in table_sheet.iter_rows(min_row=3, min_col=2, max_col=table_sheet.max_column, max_row=table_sheet.max_row):
        sum_cell = table_sheet.cell(row=row[0].row, column=row[-1].column+1)
        last_cell = table_sheet.cell(row=row[0].row, column=row[-1].column)
        sum_cell.value = f"=SUM(B{sum_cell.row}:{last_cell.column_letter}{sum_cell.row})"

    table_workbook.save(filename=get_out_filepath())
    update_debug_label("Done!", "green")


def create_frame(root, label_text, execl_path):
    load_frame = tk.Frame(master=root, relief=tk.GROOVE, borderwidth=5)
    tk.Label(master=load_frame,
             text=label_text).grid(row=1, column=1, padx=5, pady=5)
    load_label = tk.Label(text="", master=load_frame, width=30)
    load_label.grid(row=2, column=1, padx=5, pady=5)

    # This is where we lauch the file manager bar.
    def OpenFile(label):
        global EXECLS_PATHS
        name = askopenfilename(initialdir="./",
                               filetypes=(("Execl File", "*.xlsx"),
                                          ("All Files", "*.*")),
                               title="Choose a file."
                               )
        label_text = name
        if(len(name) > 30):
            label_text = "..." + name[-30:]

        label.config(text=label_text)
        EXECLS_PATHS[execl_path] = name

    tk.Button(master=load_frame, text="Browse",
              command=lambda: OpenFile(load_label)).grid(row=2, column=2, padx=5, pady=5)

    return load_frame


def tk_window():
    root = tk.Tk()
    root.title("SpreadSheeter")

    global DEBUG_LABEL
    if(DEBUG_LABEL == None):
        DEBUG_LABEL = tk.Label(
            text="Click Finish to create the SpreadSheeted table")

    DEBUG_LABEL.grid(row=4, column=1, padx=5, pady=5)

    table_frame = create_frame(
        root, "Please select the execl file to SpreadSheet: ", "Table")
    table_frame.grid(row=1, column=1, padx=20, pady=10)
    prices_frame = create_frame(
        root, "Please select the prices execl file to use: ", "Prices")
    prices_frame.grid(row=2, column=1, padx=20, pady=10)

    tk.Button(text="Finish",
              command=calculate_spreadsheet).grid(row=3, column=1, padx=20, pady=10)

    root.mainloop()


def main():
    tk_window()


if __name__ == "__main__":
    main()
