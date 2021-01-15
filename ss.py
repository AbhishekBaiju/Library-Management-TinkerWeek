import xlrd
import xlwt
from xlutils.copy import copy
path='data.xls'
rb=xlrd.open_workbook(path)
wb=copy(rb) 
w_sheet=wb.get_sheet(0)
m_sheet=wb.get_sheet(1)
inputWorkbook=xlrd.open_workbook(path)
booksheet=sheet1=inputWorkbook.sheet_by_index(0)
membersheet=sheet2=inputWorkbook.sheet_by_index(1)
#print(bookshee.nrows)
def main():
    login()
    save()
def save():
    wb.save('data.xls') 
def addbook():
    booksheet=sheet1=inputWorkbook.sheet_by_index(0)
    name=input("Enter the name of book to be added: ")
    author=input("Enter the Author's name: ")
    publisher=input("Enter the publishers name: ")
    availability=input("Enter the Availability of this book,('yes' or 'no': ")
    count=booksheet.nrows
    nextbook=count+1
    print("Total number of books is now "+str(count))
    w_sheet.write(count,0,name)
    w_sheet.write(count,1,author)
    w_sheet.write(count,2,publisher)
    w_sheet.write(count,4,publisher)
    save()
    return
def printbooks():
    books=[] 
    author=[]
    publisher=[]
    inputWorkbook=xlrd.open_workbook(path)
    booksheet=sheet1=inputWorkbook.sheet_by_index(0)
    for y in range(1,booksheet.nrows):
        books.append(booksheet.cell_value(y,0))
        author.append(booksheet.cell_value(y,1))
        publisher.append(booksheet.cell_value(y,2))
    print("\tSL.no\t\t Books\t\t Author\t\t Publisher\n")
    for i in range(len(books)):
        print("\t",i+1,end='\t')
        print("\t",books[i],end='\t')
        print("\t",author[i],end=' ')
        print("\t",publisher[i],end='\t')
        print("\n")
    print("\n\n")
    return  
def copyof():
    choice=input("Enter the name of new file (.xls)") 
    wb.save(choice+'.xls')
    return
def available():
    inputWorkbook=xlrd.open_workbook(path)
    booksheet=sheet1=inputWorkbook.sheet_by_index(0)
    booook=flag=0
    choice=input("Enter name of book: ")
    for y in range(1,booksheet.nrows):
        if (booksheet.cell_value(y,0))==choice:
            booook=y
            flag=1
            if (booksheet.cell_value(y,4))=="yes":
                print(booksheet.cell_value(y,0)+" is available!")
                break
            elif (booksheet.cell_value(y,4))=="no":
                print(booksheet.cell_value(y,0)+" is not available")
                break
            else:
                print("No value for Availability\nAdd new value\n")
                break
        else:
            print("No books named "+choice)
            break
    choice=input("\nTo change availability enter /change, else enter /0: ")
    if flag==1:
        if choice=="/change":
            qst=input("To change to 'available' enter 'yes', else enter 'no': ")
            w_sheet.write(booook,4,qst)
            save()
            return
        else:
            print("okay cool!")
            print("\n\n")
            return
def reset():
    print("Security question")
    while True:
        sqrqs=input("Name of your favorite show: ")
        if sqrqs==membersheet.cell_value(1,6):
            newhash=input("Enter new password:")
            m_sheet.write(1,5,newhash)
            save()
            print("Password Reseted!\nCountinue to login!")
            break
        else:
            print("Wrong answer\ntry again!")
def libchoices():
    inputWorkbook=xlrd.open_workbook(path)
    booksheet=sheet1=inputWorkbook.sheet_by_index(0)
    membersheet=sheet2=inputWorkbook.sheet_by_index(1)
    while True:
        print("Options for Librarian:\n")
        print("1.View all books:(enter '/all')")
        print("2.Add more books:(enter '/add')")
        print("3.Check availability or change availability of a book:(enter '/available')")
        print("4.Reset password:(enter '/reset')")
        print("5.Make a copy of xls file as new file:(enter '/copy')")
        print("6.Exit:(enter '/exit') to logout")
        choice=input("Enter choice:")
        if choice=="/all":
            print("\n")
            printbooks()
        elif choice=="/add":
            print("\n")
            addbook()
        elif choice=="/available":
            print("\n")
            available()
        elif choice=="/reset":
            reset()
        elif choice=="/copy":
            copyof()
        elif choice=="/exit":
            print("\nLogged out!")
            print("Enter '/exit' to exit!")
            break
        else:
            print("\nEnter correct choices")
def login():
        while True:
            choice=input("Librarian or user:\nEnter 1 for Librarian or 0 for User:")
            if choice=='1':
                print("Enter '/exit' to exit!")
                name=input("Enter your name: ")
                while True:
                    passs=input("Enter password to login: ")
                    hashh=membersheet.cell_value(1,5)
                    if passs==hashh:
                        print("\nWelcome back "+name+"!")
                        libchoices()
                    elif passs=='/reset':
                        reset()
                    elif passs=="/exit":
                        quit()
                    else:
                        print("Wrong password, Try again!\nEnter '/reset' for reseting password!")
            elif choice=='0':
                print("user")
                print("its being a busy week for me,\ni could finish this project as expected\ni had too many assigmnets and classes,\ni have planned a lot of funtions for this, that i could make!:(\nim so sorry for this!\nyou guys did an awsome job in teaching and i learned a lot\n so, thank you and also sorry!\n:)")
            elif choice=="/exit":
                quit()
            else:
                print("Enter choice correctly!\nenter '/exit' to stop")
if __name__ == "__main__":
    main()