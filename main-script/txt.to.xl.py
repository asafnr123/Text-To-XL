import pandas as pan
import os

# This is were the XL files will be saved to
workingDir = os.path.join(".","main-script", "MyXL")
filesExtension = ".xlsx"

# gets user input and returns a string with the file's name
def getFileName():

    # all the chars that are not allowed for an XL file
    invalidChars = ['\\', ':', '*', '?', '"', '<', '>', '|']

    
    while True:
        fileName = input("File name: ").strip()

        # checks file length and if it contains one of the invalid chars
        if 0 < len(fileName) < 260:
            if any(char in fileName for char in invalidChars):
                print('A file name can\'t contain the following letters: \\ / : * ? " < > |')
            elif fileName.startswith('.'):
                print("File can't start with .")
            else:
                # check if file name already exists
                if not os.path.exists(os.path.join(workingDir, fileName + filesExtension)):
                    # returns the file name + .xlsx
                    return fileName + filesExtension
                else:
                    print("File already exist, please choose a different name")
        else:   
            print("File must be between 1 and 259 characters")
        
           
            
            

# returns int for how many columns the user wants
def getColumnNumber():

    while True:
        columnNumbers = input("Columns number: ")

        # check if user's input is a number bigger then 0
        if columnNumbers.isdigit():
            if columnNumbers == 0:
                print("Need at least 1 column")
            else:
                return int(columnNumbers)
        else:
            print("Please enter a number")



# a loop that runs on each column and gets the user column name, returns a list with the column names
def getColumnNames(columnNumber):

    columnList = []

    for i in range(columnNumber):

        while True:
            userColumn = input(f"Column {i + 1}: ").strip().capitalize()
        
            # check if user used the same column name and if its empty
            if userColumn in columnList:
                print("Can't use the same column name twice")
            elif len(userColumn) == 0:
                print("Column can't be empty")
            else:
                columnList.append(userColumn)
                break

    return columnList


# keep asking the user for a row until he stops, returns a list of rows (each row is a list)
def getRows(columnNamesList):

    rowsList = []
    tempColumnList = []

    while True:
        for i, column in enumerate(columnNamesList):
            
            # getting user row for each current column
            userRow = input(f"{column}: ").strip()
            tempColumnList.append(userRow)

            if i == len(columnNamesList) - 1:
                
                # check if user wants a new row or stop
                while True:
                    userChoice = input("New row? (Y/N) ").strip().lower()

                    if userChoice == "y":
                        rowsList.append(tempColumnList.copy())
                        tempColumnList = []
                        break
                    elif userChoice == "n":
                        rowsList.append(tempColumnList)
                        return rowsList
                    else:
                        print("Invalid input")


# takes a list of column names and a lost of rows (each row is a list of values) returns a list of dictionaries
def formatData(columnsNameList, rowsList):

    formattedData = []

    for row in rowsList:

        tempFormattedData = {k:v for k,v in zip(columnsNameList, row)}
        formattedData.append(tempFormattedData)

    return formattedData






                
# running script
if __name__ == "__main__":

    print("Welcome to Text To XL")

    # get all the user's info
    userFileName = getFileName()
    userColumnNumber = getColumnNumber()
    userColumnNames = getColumnNames(userColumnNumber)
    userRows = getRows(userColumnNames)

    # create a pandas data frame after formatting the user's data then adding it
    df = pan.DataFrame(formatData(userColumnNames, userRows))

    # adding the file to ./main-script/file name
    try:
        df.to_excel(os.path.join(workingDir, userFileName))
    except FileNotFoundError:
        print("No such file or directory " + workingDir)
    except TypeError:
        print("Data not supported by Excel")
    except ModuleNotFoundError:
        print("openpyxl was not found. try 'pip install openpyxl' and then start again")
    except Exception as e:
        print(f"Something went wrong: {e}")

   

    





