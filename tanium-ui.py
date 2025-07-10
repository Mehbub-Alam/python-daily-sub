from tkinter import *
from tkinter import filedialog
from tkcalendar import *

root = Tk()
root.iconphoto(False, PhotoImage(file="icon.png"))

bgColor = "#1b2b36"
root.configure(background=f'{bgColor}')


screenWidth = root.winfo_screenwidth()
screenHeight = root.winfo_screenheight()

appWidth = 600
appHeight = 400

x = int((screenWidth/2) - (appWidth/2))
y = int((screenHeight/2) - (appHeight/2))

root.geometry(f'{appWidth}x{appHeight}+{x}+{y-60}')
# root.eval('tk::PlaceWindow . center')
root.title("Tanium Daily Subscription")
label = Label(root, text="Tanium Daily Subscription", font=('Arial', 18, 'bold'), bg=f'#e1edf5', fg='#ff003c', width=appWidth, height=2)
label.pack()

label2  = Label(root, text="Select Reports", font=('Arial', 14, 'bold'),bg=f'{bgColor}', fg='#e1edf5')
label2.pack(padx=20, pady = 15)

# clientStatusFile =""
# rawReportsFile = ""

def btnClick(myArgs):
    if myArgs == "clientStatus":
        global clientStatusFile
        clientStatusFile = filedialog.askopenfilename(title ="Select the Client Status file", filetypes =(("csv files", ".csv"), ("all files", ".*")))
   
    elif myArgs == "rawReport":
        global rawReportsFile
        rawReportsFile = filedialog.askopenfilename(title ="Select the Raw Report file", filetypes =(("csv files", ".csv"), ("all files", ".*")))
   
    elif myArgs == "cmdbStatus":
        global cmdbStatusFile
        cmdbStatusFile = filedialog.askopenfilename(title ="Select the CMDB File", filetypes =(("csv files", ".csv"), ("all files", ".*")))

    else:
        global nimbusSatusFile
        nimbusSatusFile = filedialog.askopenfilename(title ="Select the AWS Nimbus File", filetypes = (("csv files", ".csv"), ("all files", ".*")))

clientStatusBtn = Button(root, text ="Last Reg File", command = lambda : btnClick("clientStatus"), width=30)
clientStatusBtn.pack(pady=3)
rawReportsFileBtn = Button(root, text ="RAW Tanium Report", command = lambda : btnClick("rawReport"), width=30)
rawReportsFileBtn.pack()
cmdbStatusBtn = Button(root, text ="CMDB Report", command = lambda : btnClick("cmdbStatus"), width=30)
cmdbStatusBtn.pack(pady=3)
nimbusSatusBtn = Button(root, text ="AWS Nimbus Report", command = lambda : btnClick("nimbusSatus"), width=30)
nimbusSatusBtn.pack()

footer = Frame(root, bg=bgColor)
footer.pack(pady=15)

datePickerLabel  = Label(footer, text="Select Date", font=('Arial', 12, 'bold'),bg=f'{bgColor}', fg='#e1edf5')
datePickerLabel.pack(padx=20)

def myUpdate(*args):
    global offlineDate
    offlineDate = sel.get()

sel = StringVar()

datePicker = DateEntry(footer, date_pattern="m/d/yyyy", textvariable=sel)
datePicker.pack(pady = 15)

sel.trace("w", myUpdate)

offlineDate = datePicker.get()

# lambda *args:
submitButton = Button(root, text="Submit", bd="0", height=2, width=12, bg='#ff003c', fg="#fff", activebackground="#ba0630", activeforeground="#fff", font=('Arial', 10, 'bold'))
submitButton.pack()


# exportPath =

# root.destroy()
root.mainloop()
print("File Client Status path is: " + clientStatusFile)
print("File Raw Report path is: " + rawReportsFile)
print("File Raw Report path is: " + cmdbStatusFile)
print("File Raw Report path is: " + nimbusSatusFile)
print("Date: " + str(offlineDate))
# root = Tk()
# root.geometry("300x200")
# # root.eval('tk::PlaceWindow . center')
# root.title("Tanium Daily Subscription")
# label = Label(root, text="Running operation", font=('Arial', 18, 'bold'), bg=f'#e1edf5', fg='#ff003c', width=300, height=2)
# label.pack()
# root.mainloop()