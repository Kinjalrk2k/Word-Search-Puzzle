import xlrd

""" The following dictionary conatins all the possible directions 
    that a particular cell can follow along with their required 
    row and column updating values. Check the following syntax, for this dict:
    '<direction-name>': (<change-in-row>, <change-in-col>)
    The changes can be simply added to the index    """
directions = {
    'north': (-1, 0),
    'south': (1, 0),
    'east': (0, 1),
    'west': (0, -1),
    'north-east': (-1, 1),
    'north-west': (-1, -1),
    'south-east': (1, 1),
    'south-west': (1, -1)
}

def print_in_table(arr):
    """ This function prints a 2D array in a tabulated format   """
    numcols = len(arr[0])   #   number of cols in the passed 2D array

    print(end  = '+')       #   first row capping
    print('---+' * numcols)

    for r in arr:
        print(end = '| ')   #   head of each loop
        for c in r:
            print(c, end = ' | ')   #   after each letters
        print(end  = '\n+')     #   division between each rows
        print('---+' * numcols)

def read_excel():
    """ This function reads the puzzle.xlsx file (both the sheets)
        and returns the puzzle(2D array) and thee word querry list"""
    wb = xlrd.open_workbook('puzzle.xlsx')  #  opening the workbook to read from
    sheet = wb.sheet_by_index(0)    #   sheet object captures the first sheet

    row_offset = 2  #   for headings
    arr2d = []    #   the 2D array to store the puzzle
    for row in range(sheet.nrows - row_offset):
        temp_row = []
        for col in range(sheet.ncols):
            temp_row.append(sheet.cell(row + row_offset, col).value)
        arr2d.append(temp_row)

    sheet = wb.sheet_by_index(1)    #   sheet object captures the second sheet

    arr1d = []  # the 1D array to store the query words
    for row in range(sheet.nrows - row_offset):
        arr1d.append(sheet.cell(row + row_offset, 0).value)

    return (arr2d, arr1d)   #   return both the puzzle and word list


puzzle, words  = read_excel()  #   this is the 2D array containing the puzzle
print_in_table(puzzle)
print(words)