from openpyxl import Workbook, load_workbook
import os.path, sheet_handler

# TO DO: internally rename to sheet_menu

# Simple menu. Has some input handling to ensure proper inputs.
def onStart():
    selection = 0
    while selection not in [1,2]:
        print("\nSelect from the following: "+
          "  \n  1. Add/Remove cards" +
          "  \n  2. Exit")
        try: 
            selection = int(input("\nSelection: "))
        except ValueError:
            print("Please select from option 1 or 2.")
            continue            
        else:
            if selection in [1,2]:
                selection = selectionHandler(selection)
                if selection == 2:
                    break
            else:
                print("Please select from option 1 or 2.")
                continue

def selectionHandler(selection):
    match int(selection):
        case 1:
            print("\nMoving to add menu...")
            workbook = sheet_handler.createWorkbook()
            addCardMenu(workbook)
            return 0
        case 2:
            print("\nExiting back to main menu...")
            return 2



def addCardMenu(workbook):
    sheet_handler.addCards(workbook)
    
