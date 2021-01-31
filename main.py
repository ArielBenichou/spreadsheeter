from openpyxl import Workbook


def alphanum_generator(from_num, to_num):
    """ yield an xlsx column format from number up to another number
        from_num must be greater than 1
        1 will yield A (chr(65) == "A" )
        26 will yield Z ( hr(90) == "Z")
    """
    if(from_num <= 0):
        raise ValueError
    counter = 0
    while counter <= to_num:
        yield chr(64+from_num+counter)


def create_test_workbook():
    workbook = Workbook()
    sheet = workbook.active
    food_dic = {"1111": "cumcumber", "1112": "tomato",
                "1114": "pepper", "2222": "milk", "3142": "meat"}

    # init the first two row with names and ids

    for i in range(len(food_dic)):
        sheet["B"+i] = "hello"
        sheet["C"+i] = "world!"

    workbook.save(filename="hello_world.xlsx")


def main():
    pass


if __name__ == "__main__":
    main()
