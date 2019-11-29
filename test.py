from tkinter import *
from openpyxl import load_workbook
from tkinter import messagebox
from PIL import ImageTk, Image


def end(event):
    root.destroy()


def print_check():
    eName = entry_1.get()
    eDate = entry_2.get()
    eItem = entry_3.get()
    eCost = entry_4.get()
    eProject = entry_5.get()
    eSupplier = entry_6.get()

    workbook_name = "Budgeting.xlsx"
    wb = load_workbook(workbook_name)
    page = wb.active

    all_entries = [[eName, eDate, eItem, eCost, eProject, eSupplier]]

    if len(eName) == 0:
        messagebox.showerror("Error", "Please Enter Your Name")
    if len(eDate) == 0:
        messagebox.showerror("Error", "Please Enter The Date")
    if len(eItem) == 0:
        messagebox.showerror("Error", "Please Enter The Item Purchased")
    if len(eCost) == 0:
        messagebox.showerror("Error", "Please Enter The Cost")
    if len(eProject) == 0:
        messagebox.showerror("Error", "Please Enter Your Current Project")
    if len(eSupplier) == 0:
        messagebox.showerror("Error", "Please Enter The Supplier")
    if len(eName) and len(eDate) and len(eItem) and len(eCost) and len(eProject) and len(eSupplier) >= 1:
        for info in all_entries:
            page.append(info)
    wb.save(filename=workbook_name)

    root.destroy()


root = Tk()

root.title("Budgeting")

root.geometry("600x380")

'#top menu frame'
menuFrame_1 = Frame(root, bg="#060273", height="30")
menuFrame_1.pack(fill=X)

'#make space'
spaceFrame_1 = Frame(root, height="5")
spaceFrame_1.pack(fill=X)

'#title frame'
tFrame_1 = Frame(root)
title_1 = Label(tFrame_1, text="Enter Information")
title_1.config(font=("Arial", 18))
title_1.pack()
tFrame_1.pack()

'#main frame'
mainFrame_1 = Frame(root, borderwidth="2", relief=SUNKEN, width="400", height="180")
mainFrame_1.pack(fill=None, expand=False)

'#label frame'
lFrame_1 = Frame(mainFrame_1)
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
lFrame_1.pack(side=LEFT)

'#entry box frame'
eFrame_1 = Frame(mainFrame_1)
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
eFrame_1.pack(side=RIGHT)

'#make space'
spaceFrame_2 = Frame(root, height="5")
spaceFrame_2.pack(fill=X)

'#button frame'
bFrame_1 = Frame(root)
enter_1 = Button(bFrame_1, text="Enter", command=print_check, width=12, height=2)
enter_1.config(font=("Arial", 18))
enter_1.grid(row=0, columnspan=4)
bFrame_1.pack()

'#make space'
spaceFrame_3 = Frame(root, height="5")
spaceFrame_3.pack(fill=X)

'#bottom menu frame'
bMenuFrame_1 = Frame(root, bg="#060273", height="30")
canvas_1 = Canvas(bMenuFrame_1, height="30", width="570", bg="#060273", highlightthickness=0)
powerButton = Image.open("power3.png")
powerButton = powerButton.resize((20, 20), Image.ANTIALIAS)
myimg = ImageTk.PhotoImage(powerButton)
canvas_1.create_image(15, 15, image=myimg)
canvas_1.bind("<Button>", end)
canvas_1.pack(side=LEFT)
bMenuFrame_1.pack(fill=X)

root.mainloop()
