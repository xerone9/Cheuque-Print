from tkinter import *
from tkinter import filedialog
from askari_Chq import askari_Chq
from askari_Chq import alhabib_Chq
import tkinter.font as font


def UploadAction():
    filetypes = (
        ('Excel files', '*.xlsx'),
    )

    filename = filedialog.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    file = open("text.txt", "w")
    file.write(filename)
    file.close()
    file = open("text.txt", "r")
    data = file.read()
    getLocation.config(text=data, font=(8))
    savedLocation.config(text=data)


def askariPrint():
    f = open("text.txt", "r")
    data = f.read()
    askari_Chq(data)


def alhabibPrint():
    f = open("text.txt", "r")
    data = f.read()
    alhabib_Chq(data)

f = open("text.txt", "r")
data = f.read()
root = Tk()
root.resizable(0,0)
root.title('Cheque Printing Software (Beta) - V-1.0')
root.geometry("400x340")
root.configure(bg="white")


# myFont = font.Font(family='Helvetica', size=20, weight='bold')

label = Label(root, text="Cheque Printing Software", font=("French Script MT", 35, 'bold'))
label.configure(foreground="blue")
label.configure(bg="white")
label.place(x=10, y=20)

getLocation = Label(root, text="Select Excel File (.xlsx only)", font=(10))
getLocation.configure(bg="white")
getLocation.place(x=18, y=95)

browseButton = Button(root, text='Browse', command=UploadAction)
browseButton.place(x=325, y=95)

savedLocationLabel = Label(root, text='Saved Location: ', font=("Helvetica", 11, 'bold'))
savedLocationLabel.configure(bg="white")
savedLocationLabel.place(x=26, y=135)

savedLocation = Label(root, text=data)
savedLocation.configure(bg="white")
savedLocation.place(x=145, y=137)

bankLabel = Label(root, text="* Select Bank Template *", font=("Arial", 19, 'bold'))
bankLabel.configure(foreground="brown")
bankLabel.configure(bg="white")
bankLabel.place(x=55, y=180)

imgAskari = PhotoImage(file='askari.png')
askariButton = Button(root, image=imgAskari, text="Askari", command=askariPrint)
askariButton.place(x=20, y=237)

imgAlhabib = PhotoImage(file='alhabib.png')
alHabibButton = Button(root, image=imgAlhabib, text="Al-Habib", command=alhabibPrint)
alHabibButton.place(x=215, y=237)

footer = Label(root, text="                By - Usman Mustafa Khawar                 ", font=(14))
footer.configure(foreground="white")
footer.configure(bg="black")
footer.pack(side=BOTTOM)


root.mainloop()
