# SpreadSheeter

This is a script to help my father in his job.

The program mulitpy each column in each sheet by the number in the first cell of the column, if the cell is empty we skip it.

We add a sum cell at the end of each row.

We take all the sum cell and make a sum table in a seperate sheet.

## Install

Install `python3`, and then install the `openpyxl` module. (with `pip install openpyxl` in the terminal)

DONE! now you can run the script - from the terminal/cmd or double click it.

If you want you can compile the program with `pyinstaller`:

- install the module - `pip install pyinstaller`
- run the command (be in the dir of the script) - `pyinstaller --onefile spreadsheeter.py`

## Useage

The program open up with a simple UI: a 'Browse' button and one 'Finish' button.
What to do:

1. load the table with the first button.
2. click finish!
3. in the same folder of the first table you have an "\_out.xlsx" file.
4. DONE!

### Useful to know

- The program skips rows that don't have something in the first cell of a row. (AKA each row need to have a client)
