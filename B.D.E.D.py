from tkinter import *
from openpyxl import load_workbook
from PIL import ImageTk, Image
from tkinter import messagebox


def help_window(event):

    def close_window():
        root_2.destroy()

    root_2 = Tk()

    root_2.title("B.D.E.D")

    root_2.geometry("600x380")

    root_2.resizable(False, False)
    
    '#Creates the top menu frame on root_2'
    menuFrame_2 = Frame(root_2, bg="#060273", height="30")
    menuFrame_2.pack(fill=X)

    '#Makes space on root_2'
    spaceFrame_4 = Frame(root_2, height="15")
    spaceFrame_4.pack(fill=X)
    
    '#Creates title frame on root_2' 
    tFrame_2 = Frame(root_2)
    tFrame_2.pack()

    '#Creates title frame on root_2'
    title_2 = Label(tFrame_2, text="Help")
    title_2.config(font=("Arial", 18))
    title_2.pack()
    
    '#Makes space on roo_2'
    spaceFrame_5 = Frame(root_2, height="10")
    spaceFrame_5.pack(fill=X)

    '#Creates the main frame to put label widgets on for root_2'
    mainFrame_2 = Frame(root_2, borderwidth="2", relief=SUNKEN, width="400", height="180")
    mainFrame_2.pack()

    '#Create and add label widgets to root_2'
    label_7 = Label(mainFrame_2, text="When Entering Yor Name Use Your Full Name")
    label_8 = Label(mainFrame_2, text="Eg: John Doe", fg="#060273")
    label_9 = Label(mainFrame_2, text="When Entering The Date Enter It In The Following Format: Month, Day, Year")
    label_10 = Label(mainFrame_2, text="Eg: December, 1, 2019", fg="#060273")
    label_11 = Label(mainFrame_2, text="When Entering The Cost You Don't Have To Include A Dollar Sign ")
    label_12 = Label(mainFrame_2, text="Eg: 299.99", fg="#060273")
    label_13 = Label(mainFrame_2, text="Make Sure To#060273 Enter All Details Properly")
    label_14 = Label(mainFrame_2, text="For Any Further Questions Contact matt.lipton@ucc.on.ca", fg="#060273")
    label_7.grid(row=0)
    label_8.grid(row=1)
    label_9.grid(row=2)
    label_10.grid(row=3)
    label_11.grid(row=4)
    label_12.grid(row=5)
    label_13.grid(row=6)
    label_14.grid(row=7)
    
    '#Makes space on root_2'
    spaceFrame_6 = Frame(root_2, height="10")
    spaceFrame_6.pack(fill=X)

    '#Creates the button frame on root_2'
    bFrame_2 = Frame(root_2, height="20")
    bFrame_2.pack()
    
    '#Creates enter button that is closes the window'
    enter_2 = Button(bFrame_2, text="Close Window", command=close_window, width=12, height=2)
    enter_2.config(font=("Arial", 18))
    enter_2.grid(row=0, columnspan=4)
    
    '#Creates bottom menu frame on root_2'
    bMenuFrame_2 = Frame(root_2, bg="#060273", height="30")
    bMenuFrame_2.pack(fill=X, side=BOTTOM)

    root_2.mainloop()


def end(event):
    root.destroy()


def print_check():
    eN = entry_1.get()
    eD = entry_2.get()
    eI = entry_3.get()
    eC = entry_4.get()
    eP = entry_5.get()
    eS = entry_6.get()

    '#Opens B.D.E.D.xlsx'
    workbook_name = "B.D.E.D.xlsx"
    wb = load_workbook(workbook_name)
    
    '#Sets page variable to allow editing of B.D.E.D.xlsx'
    page = wb.active
    
    '#Creates list of all the entry box data'
    all_entries = [[eN, eD, eI, eC, eP, eS]]

    '#Error checking for each entries length and for cost if it is a number or not'
    if eC.isalpha():
        messagebox.showerror("Error", "Please Enter The Cost Properly", icon="warning")
    if len(eN) == 0:
        messagebox.showerror("Error", "Please Enter Your Name", icon="warning")
    if len(eD) == 0:
        messagebox.showerror("Error", "Please Enter The Date", icon="warning")
    if len(eI) == 0:
        messagebox.showerror("Error", "Please Enter The Item Purchased", icon="warning")
    if len(eC) == 0:
        messagebox.showerror("Error", "Please Enter The Cost", icon="warning")
    if len(eP) == 0:
        messagebox.showerror("Error", "Please Enter Your Current Project", icon="warning")
    if len(eS) == 0:
        messagebox.showerror("Error", "Please Enter The Supplier", icon="warning")
        
    '#Checks if all the data was entered properly then adds it to B.D.E.D.xlsx'
    if len(eN) and len(eD) and len(eI) and len(eC) and len(eP) and len(eS) >= 1 and eC.isdigit():
        ayn = messagebox.askquestion("B.D.E.D", "Are You Sure")
        if ayn == 'yes':
            for info in all_entries:
                page.append(info)
            '#Saves the file then closes the window to prevent duplicate entries'
            wb.save(filename=workbook_name)
            root.destroy()


root = Tk()

root.title("B.D.E.D")

root.geometry("600x380")

root.resizable(False, False)

'#Creates the top menu frame'
menuFrame_1 = Frame(root, bg="#060273", height="30")
menuFrame_1.pack(fill=X)

'#Makes space on GUI'
spaceFrame_1 = Frame(root, height="5")
spaceFrame_1.pack(fill=X)

'#Creates title frame'
tFrame_1 = Frame(root)
tFrame_1.pack()

'#Creates and adds a title to the GUI'
title_1 = Label(tFrame_1, text="Enter Information")
title_1.config(font=("Arial", 18))
title_1.pack()


'#Creates the main frame to put label and entry box frame on'
mainFrame_1 = Frame(root, borderwidth="2", relief=SUNKEN, width="400", height="180")
mainFrame_1.pack(fill=None, expand=False)

'#Creates label frame'
lFrame_1 = Frame(mainFrame_1)
lFrame_1.pack(side=LEFT)

'#Create and add label widgets to GUI'
label_1 = Label(lFrame_1, text="Name", bg="white", fg="black", height=2)
label_2 = Label(lFrame_1, text="Date of Order", bg="white", fg="black", height=2)
label_3 = Label(lFrame_1, text="Item Purchased", bg="white", fg="black", height=2)
label_4 = Label(lFrame_1, text="Cost of Purchase", bg="white", fg="black", height=2)
label_5 = Label(lFrame_1, text="Project", bg="white", fg="black", height=2)
label_6 = Label(lFrame_1, text="Supplier", bg="white", fg="black", height=2)
label_1.grid(row=1, sticky=W)
label_2.grid(row=3, sticky=W)
label_3.grid(row=5, sticky=W)
label_4.grid(row=6, sticky=W)
label_5.grid(row=7, sticky=W)
label_6.grid(row=9, sticky=W)

'#Creates entry box frame'
eFrame_1 = Frame(mainFrame_1)
eFrame_1.pack(side=RIGHT)

'#Create and add entry box widgets to GUI'
entry_1 = Entry(eFrame_1)
entry_2 = Entry(eFrame_1)
entry_3 = Entry(eFrame_1)
entry_4 = Entry(eFrame_1)
entry_5 = Entry(eFrame_1)
entry_6 = Entry(eFrame_1)
entry_1.grid(row=1, ipady=5)
entry_2.grid(row=2, ipady=5)
entry_3.grid(row=3, ipady=5)
entry_4.grid(row=4, ipady=5)
entry_5.grid(row=5, ipady=5)
entry_6.grid(row=6, ipady=5)

'#Makes space on GUI'
spaceFrame_2 = Frame(root, height="5")
spaceFrame_2.pack(fill=X)

'#Creates the button frame'
bFrame_1 = Frame(root)
bFrame_1.pack()

'#Creates enter button that is bound to print_check'
enter_1 = Button(bFrame_1, text="Enter", command=print_check, width=12, height=2)
enter_1.config(font=("Arial", 18))
enter_1.grid(row=0, columnspan=4)

'#Makes space on GUI'
spaceFrame_3 = Frame(root, height="5")
spaceFrame_3.pack(fill=X)

'#Creates bottom menu frame'
bMenuFrame_1 = Frame(root, bg="#060273", height="30")
bMenuFrame_1.pack(fill=X)

'#Creates first canvas on bottom menu frame and adds power_3.png to it'
canvas_1 = Canvas(bMenuFrame_1, height="30", width="570", bg="#060273", highlightthickness=0)
powerButton = Image.open("power3.png")
powerButton = powerButton.resize((20, 20), Image.ANTIALIAS)
myimg = ImageTk.PhotoImage(powerButton)
canvas_1.create_image(15, 15, image=myimg)
canvas_1.bind("<Button>", end)
canvas_1.pack(side=LEFT)

'#Creates second canvas on bottom menu frame and adds help_button.png to it'
canvas_2 = Canvas(bMenuFrame_1, height="30", bg="#060273", highlightthickness=0)
questionButton = Image.open("help_button.png")
questionButton = questionButton.resize((20, 20), Image.ANTIALIAS)
myimg_2 = ImageTk.PhotoImage(questionButton)
canvas_2.create_image(15, 15, image=myimg_2)
canvas_2.bind("<Button>", help_window)
canvas_2.pack(side=RIGHT)

root.mainloop()
