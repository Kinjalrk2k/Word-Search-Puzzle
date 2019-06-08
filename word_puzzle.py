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

    row_offset = 0  #   for headings
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

def excelcol(num):
    """ This function converts an index number to excel styled column letter"""
    ecol = ""
    while num > 0:
        rem = (num - 1) % 26    
        num = (num - 1) // 26
        ecol = ecol + chr(65 + rem)

    return ecol


puzzle, words_raw  = read_excel()  #   unpaking tupples to puzzle and word

print('Puzzle')
print_in_table(puzzle)

words1 = [x.upper() for x in words_raw]  #   converting all words to uppercase
words2 = [x.replace(' ', '') for x in words1]   #   removing all spaces
words = words2

print(f'Queried Words: {words}')

numrows = len(puzzle)       #   no. of rows in the puzzle
numcols = len(puzzle[0])    #   no. of cols in the puzzle

poss_words = {}     #   list of possible words in all directions for a particular letter (cell)
temp_word = []      #   temp list of chars of the words

for w in words:     #   iterate for each queried word
    found = False

    #   iterating over the 2D array
    for r in range(numrows):
        for c in range(numcols):

            # dname is the key(direction name) and dupdate is the value(update values)
            for dname, dupdate in directions.items():
                #   wought to be the new indicies
                newr = r
                newc = c
                i, j = dupdate  #   unpack tupple for seperate row and col updation

                 #   constrain for gatthering letters till the length of the queried word
                for l in range(len(w)):
                    temp_word.append(puzzle[newr][newc])    #   append the letters one-by-one
                    # update indices
                    newr += i   
                    newc += j
                    if newr < 0 or newr >= numrows or newc < 0 or newc >=numcols:   #   added wrapping constraints
                        break

                string = ''.join(temp_word)     #   converting the char list to string
                poss_words.update({dname : string}) #   adding new possible word along with direction key
                temp_word.clear()   #   clear for next iteration

            #   check if the quried word is present in the possible words dict
            for p_word_dir, p_word in poss_words.items():
                if w == p_word:
                    found  = True
                    #   print the word, along with the staring letter postion in the puzzle and direction
                    print(f'Word: {w},\t\t Postion: {excelcol(c+1)}{r+1},\t Direction: {p_word_dir}')
                    break

            poss_words.clear()  #   clear for next iteration
            
            if found is True:   #   already found, no need to iterate more
                break

        if found is True:   #   already found, no need to iterate more
            break
    
    if found is False:  #   when the word is not found
        print(f'Word: {w},\t\t NOT FOUND')