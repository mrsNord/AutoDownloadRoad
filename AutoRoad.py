# AutoROAD.py
# Downloads Road plans from the ROAD Archived Documents (ROAD) 

from tkinter import filedialog
from tkinter import ttk
from tkinter import *
import glob

# import libraries for AutoSticker
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlrd
import os
import urllib.request
from os.path import basename
import posixpath

searchURL = 'https://road.azdot.gov/search/state'
count = 1

# get Chrome options
chrome_options = Options()
chrome_options.add_argument("--kiosk-printing")

#main window
root = Tk()
root.title("AutoROAD")
MainFrame = Frame(root)
MainFrame.pack()

# title
roottext = Label(MainFrame, text = "AutoROAD", font=("Arial bold", 44), fg = "#91111B")
roottext.pack(side = RIGHT, padx = 25)

# dataBot
canvas = Canvas(MainFrame, width = 180, height = 213)          
img = PhotoImage(file= os.getcwd() + '\\dataBot.png')      
canvas.create_image(20,20, anchor = NW, image=img) 
canvas.pack(side = RIGHT)

#### Select File Frame ####
fileframe = LabelFrame(root, text = "Select File")
fileframe.pack(pady = 5, padx = 5, ipadx = 4)

# Filename Text
file_label = Label(fileframe, text = "Filename:")
file_label.grid(row = 1, column = 0, pady = 5, padx = 5)

# Filename Entry Box
fsr = StringVar() # only need stringvar to set entry boxes.
e_filename = Entry(fileframe, textvariable=fsr, font=("Calibri"), width = 22)
e_filename.grid(row = 1, column = 1, columnspan=2, pady = 5, padx = 5, ipadx = 52)

# Browse for Excel File
def getFile():
    fsr.set(filedialog.askopenfilename(initialdir = "/", title = "Select file", filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))) # set the text box to the input from file dialog.

browsebutton = Button(fileframe, text = "Browse", fg = "black", command = getFile)
browsebutton.grid(row = 2, column = 0, pady = 5, padx = 5)

# Select User Text
techLabel = Label(fileframe, text = "Select User:")
techLabel.grid(row = 2, column = 1)

# Select User Box
techName = StringVar(root)
techName = ttk.Combobox(fileframe, font=("Calibri"), values = ["Jane", "Celeste", "Sarah", "Gaby", "Daniel"])
techName.grid(row = 2, column = 2)


#------------------------------------------------------------------------------------------------------------------
# Select Sheet
sheetName = 'Project List'
#------------------------------------------------------------------------------------------------------------------

# Secondary Frame
SecondaryFrame = Frame(root)
SecondaryFrame.pack()

# Options Frame
optionsLF = LabelFrame(SecondaryFrame, text="Options")
optionsLF.pack(side = LEFT, ipady = 37, ipadx = 5)

# Options
CheckVar1 = IntVar(value = 0)
CheckVar2 = IntVar(value = 0)
CheckVar3 = IntVar(value = 0)
C1 = Checkbutton(optionsLF, text = "Assign DTR Number", variable = CheckVar1, onvalue = 1, offvalue = 0)
C2 = Checkbutton(optionsLF, text = "Print Confirmation Sheet", variable = CheckVar2, onvalue = 1, offvalue = 0)
C3 = Checkbutton(optionsLF, text = "Print DTR Instructions", variable = CheckVar3, onvalue = 1, offvalue = 0)
C1.grid(row = 0, column = 0, sticky = 'W', padx = 5, pady = 2)
C2.grid(row = 1, column = 0, sticky = 'W', padx = 5, pady = 2)
C3.grid(row = 2, column = 0, sticky = 'W', padx = 5, pady = 2)

def runGo():
    global count, ROADdata
    
    driver = webdriver.Chrome(os.getcwd() + '\chromedriver.exe', options=chrome_options)

    # import ROAD data from excel
    ROADdata = pd.read_excel (e_filename.get(), sheet_name=sheetName, converters = {'Tracs':str}) 
    df = pd.DataFrame(ROADdata, columns = ['Tracs', 'Download'])

    # access ROAD
    driver.get(searchURL)
    time.sleep(5) # wait

    iAgree_button = driver.find_element_by_xpath('//*[@id="termsofUse"]/div/div/div[3]/a')
    iAgree_button.click()
    time.sleep(3) # wait

    # navigate to TRACS look up page
    driver.get(searchURL)
    time.sleep(5)

    def getdtr(df):
        global count
        # search for plan by TRACS number
        TRACS_box = driver.find_element_by_class_name('select2-search__field')
        TRACS_box.send_keys(df[0])
        time.sleep(18)
        TRACS_box.send_keys(Keys.RETURN)
        time.sleep(5)
        submitButton = driver.find_element_by_class_name('btn.btn-primary')
        submitButton.click()

        time.sleep(15)

        docLocation = driver.find_element_by_xpath("//tbody[contains(concat(' ',normalize-space(@class),' '),' adot-mvcDataGrid-tbody ')]/tr/td[1]//a")
        docName = driver.find_element_by_xpath("//tbody[contains(concat(' ',normalize-space(@class),' '),' adot-mvcDataGrid-tbody ')]/tr/td[1]//a/span")

        print('Downloading Document: ' + df[0])

        print(docName.getText())
        urllib.request.urlretrieve(docLocation.get_attribute("href"), os.getcwd() + '\\Docs\\' + docName.getText() + '.pdf')

        time.sleep(30)

        driver.get(searchURL)
        time.sleep(5)


    # apply the getdtr function to Data Frame
    df.apply(getdtr, axis = 1)

    driver.close()
    print('done')

# GO!!
sticker = PhotoImage(file = os.getcwd() + '\\stickGo.png')
stickerImg = sticker.subsample(5, 5) 

goButton = Button(SecondaryFrame, image = stickerImg,  command = runGo)
goButton.pack(side = LEFT, pady = 10, padx = 9, ipadx = 2, ipady = 0)

root.mainloop()
