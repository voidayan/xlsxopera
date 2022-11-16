from xlsxopera.Notebook import Notebook

book_1 = Notebook("tests/test.xlsx")
book_2 = Notebook("tests/test.xlsx", "rows and cols")
book_3 = Notebook("tests/test.xlsx", "math")

## TESTS
# Create pyTests for below --> Example data in test.xlsx should be a documentation with description
# W pliku xlsx wszystkie przykladowe dane to mogą byc opisy i dokumentacja parametrow (dobry pomysl)

print(f"list_rows (sheet rows and cols) :: {book_2.list_rows()}")
print(f"list_rows (sheet math) :: {book_3.list_rows()}")
#
print(f"dict_rows (sheet math) :: {book_3.dict_rows()}")
#
print(f"list_cols (default sheet) :: {book_1.list_cols()}")
print(f"list_cols (sheet rows and cols) from->to :: {book_2.list_cols(start=3, end=6)}")
print(f"list_cols (sheet math) from->to :: {book_3.list_cols(start=1, end=4)}")
#
print(f"dict_cols (sheet math) :: {book_3.dict_cols()}")
#
print(f"data_convert :: {book_2.data_convert()}")
#
print(f"header_values :: {book_2.header_values('DATE')}")
print(f"header_values :: {book_2.header_values('NOTES')}")
#
print(f"find :: {book_2.find('Task-1')}")
print(f"find :: {book_2.find('Task-4')}")
#
print(f"value_position :: {book_2.value_position()}")
print(f"value_position :: {book_2.value_position('PROJECT-7')}")
print(f"value_position :: {book_2.value_position('Task-5')}")
#
print(f"value_coordinates :: {book_2.value_coordinates('PROJECT-9')}")
print(f"value_coordinates :: {book_2.value_coordinates('Task-2')}")
print(f"value_coordinates :: {book_2.value_coordinates('Task-2')}")


####### TODO:
# zmiana nazwy pliku na lowercase i nazwy katalogu na lowercase
# dodawanie cells if int from cell to cell
# odejmowanie
# mnozenie
# get cell color
# get cell info: color, font, pogrubienie, podkreslenie, itp itd i wszystko return Dict np. "color": FFF12313"
# write info to cell : znajdz cell z value i zastąp go nowym value
## ORGANIZE ##
# strona internetowa z dokumentacja i reklama
