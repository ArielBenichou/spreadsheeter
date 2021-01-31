from openpyxl import Workbook, load_workbook
from openpyxl.utils import FORMULAE
import json
import sys


FOOD_CONFIG_PATH = "food_config.json"
LOAD_WORKBOOK_PATH = "table.xlsx"
OUT_WORKBOOK_SUFFIX = "_out"

if len(sys.argv) > 1:
    LOAD_WORKBOOK_PATH = sys.argv[1]


def get_out_filepath():
    prefix, delimiter, postfix = LOAD_WORKBOOK_PATH.rpartition(".")
    return prefix + OUT_WORKBOOK_SUFFIX + delimiter + postfix


def load_food_config():
    with open(FOOD_CONFIG_PATH, "r") as f:
        return json.load(f)


def get_price_of_food_id(jf, id):
    for food in jf["foods"]:
        if(str(id) == food["id"]):
            return food["price"]
    raise ValueError(f"There is not a food id of: {id}")


def main():
    jf = load_food_config()
    workbook = load_workbook(filename=LOAD_WORKBOOK_PATH)
    sheet = workbook.active

    for col in sheet.iter_cols(min_row=1, min_col=2, max_row=sheet.max_row, max_col=sheet.max_column):
        col_food_price = get_price_of_food_id(jf, col[0].value)
        for cell in col:
            if cell.row == 1 or cell.row == 2:
                continue
            cell_value = cell.value if cell.value else 0
            cell.value = float(cell_value)*float(col_food_price)

    for row in sheet.iter_rows(min_row=3, min_col=2, max_col=sheet.max_column, max_row=sheet.max_row):
        sum_cell = sheet.cell(row=row[0].row, column=row[-1].column+1)
        last_cell = sheet.cell(row=row[0].row, column=row[-1].column)
        sum_cell.value = f"=SUM(B{sum_cell.row}:{last_cell.column_letter}{sum_cell.row})"

    workbook.save(filename=get_out_filepath())


if __name__ == "__main__":
    main()