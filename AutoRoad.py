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

count = 1

# get Chrome options
chrome_options = Options()
chrome_options.add_argument("--kiosk-printing")

#main window
root = Tk()
root.title("AutoROAD")
MainFrame = Frame(root)
MainFrame.pack()
