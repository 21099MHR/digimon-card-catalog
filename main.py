import sheet_edit, sheet_sort

def menu():
    print("Select from the following: "+
          "  \n  1. Add or remove cards" +
          "  \n  2. Sort cards" +
          "  \n  3. Exit")
    selection = 0
    while selection not in [1,2,3]:
        try: 
            selection = int(input("\nSelection: "))
        except ValueError:
            print("Please select from option 1, 2, or 3.")
            continue            
        else:
            if selection in [1,2,3]:
                break
            else:
                print("Please select from option 1, 2, or 3.")
                continue
        
    match int(selection):
        case 1:
            print("\nMoving to edit menu...")
            sheet_edit.onStart()
            menu()
        case 2:
            print("\nThis feature is still being worked on! Sorry!")
            # print("\nMoving to sort menu...")
            # sheet_sort.onStart()
            menu()
        case 3:
            print("\nExiting...")

menu()
