import xlwings as xw
import sys

if len(sys.argv) != 3:
  print ("usage: listToExcel.py inputFile outputFile")
  sys.exit()

inputFile = sys.argv[1]
outputFile = sys.argv[2]

rowsPerPage = 50
columnsPerPage = 4
firstColumn = 'A'
firstRow = '1'
fileLength = len(open(inputFile).readlines())

#Pagescomplete goes from 0 to (file length/(columnsPerPage*rowsPerPage))
pagesComplete = 0;

#Spacing between consecutive pages
pageSpacing = 3;

#Curr column goes from 0 to columnsPerPage
currColumn = 0;

#Total column number goes from 0 to (file length / 50) - 1
totalColumns = 0;

#Curr row goes from 0 to rowsPerPage
currRow = 0;

wb = xw.Book(outputFile)
readFile = open(inputFile, "r")

#Iterate over pages of 200 words each
while pagesComplete < (fileLength/(rowsPerPage * columnsPerPage)):
  #Create a row offset
  offset = pagesComplete * (rowsPerPage + pageSpacing)
  currColumn = 0
  #Iterate over columns in a single page
  while currColumn < columnsPerPage:
    currRow = 0
    #Iterate over rows in a single column
    while currRow < rowsPerPage:
      #Set the number
      xw.Range(chr(ord(firstColumn) + 2 * currColumn) \
      + str(currRow + 1 + offset)).value \
      = (currRow + 1) + rowsPerPage * totalColumns
      
      #Set the value of the string
      xw.Range(chr(ord(firstColumn) + 2 * currColumn + 1) \
      + str(currRow + 1 + offset)).value \
      = readFile.readline().rstrip()
      currRow += 1
    currColumn+=1
    totalColumns+=1
  pagesComplete += 1
