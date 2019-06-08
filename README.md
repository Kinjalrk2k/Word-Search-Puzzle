# Word-Search-Puzzle
Its really a tedious job to find small words from a large NxN puzzle full of letters. So, this program attempts to ease of the pain, quickly picking up the required words!

### XLRD Module
This project uses xlrd module to read data from MSExcel file (.xlsx extension)

To install xlrd module, use:
````pip install xlrd````

or visit PyPi site, by clicking [here](https://pypi.org/project/xlrd/)

### Working
````puzzle.xlsx```` file should be in the same directory with the ````word_puzzle.py```` pyhton script

* Open the ````puzzle.xlsx```` file. It has two Sheets
    * **Puzzle** Sheet should contain the puzzle with letters. It should Begin from A1, spreading horizontally and vertically.
    * **Words to find** Sheet should contain all the words that are needed to be found from the puzzle. Again, it should begin from A1 and spread horizontally only. *(A1, A2, A3, ....., An)*
* Run the word_puzzle.py python script in the working directory. The output will be printed in the terminal.

### Reading the Output
Output has a general syntax:
```` Word:<queried_word>, Position:<cell_position>, Direction:<compass_direction> ````
* ````<queried_word>````: Words that are needed to be found, in a sequenced list.
* ````<cell_position>````: Excel styled cell ID holding the first Letter of the queried word.
* ````<compass_direction>````: Direction to look for, following the first letter to find the rest of the word. Follow the diagram below, for reference:

````
  NW N NE
   \ | /
W -- * -- E
   / | \
  SW S SE
````

### Notes
This project was created and tested under Windows.

This project is still under development. Parts of the source codes may not be well documented.
Also suitable prompts may not be available for the user at the moment.

More features and fixes are yet to come. Meanwhile suggestions, ideas, bug reports are welcomed.

<br>***Kinjal Raykarmakar***