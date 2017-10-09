
import os
from tkinter import filedialog
from Tkinter import *
import skimage.io as io
import xlsxwriter
import numpy as np
import ntpath
import tkMessageBox
import ttk

root = Tk()
root.geometry("300x300")
topFrame = Frame(root)
topFrame.pack()
bottomFrame = Frame(root)
progressbar = ttk.Progressbar(bottomFrame, orient=HORIZONTAL, length=100, mode='indeterminate')
progressbar.pack(side="bottom")


def main():
    bottomFrame.pack(side=BOTTOM)
    button1 = Button(topFrame, text="Select  Folder ", fg="red" , command = directory_open)  # fg foreground is optional
    button1.pack(side=LEFT, fill=X)
    button2 = Button(topFrame, text="Select File  ", fg="red", command=file_open)  #
    button2.pack(side=LEFT, fill=X)
    root.mainloop()

def directory_open():
    dirname = filedialog.askdirectory( initialdir="/", title='Please select a directory')

    if len(dirname) > 0:
        print "You chose %s" % dirname
        iterateDirectory(dirname)
        tkMessageBox.showwarning(
            "Complete",
            "All file are processed  and saved complete "
        )

    else:
        print 'Select Nothing !!'



def file_open():
    filename = filedialog.askopenfile(initialdir="/", filetypes=[('tif', '*.tif'),
                                                            ('tiff', '*.tiff')], title='Select File')

    if filename != None:
        initOperations(filename.name)
        tkMessageBox.showwarning(
            "Complete",
            "File is processed  and saved complete "
        )
    else:
        print 'Select Nothing !!'


def iterateDirectory(dirname):
    for filename in os.listdir(dirname):
        pathname = os.path.join(dirname, filename)
        print "Process  %s" % pathname
        if filename.endswith(".tif") or filename.endswith(".tiff"):  #process file
            initOperations(pathname)
            continue
        elif os.path.isdir(pathname): #process directory
            print "It is dir => %s" % pathname
            iterateDirectory(pathname)
            continue
        else:
            continue


def initOperations(filename):
    print 'initOperations=>' + filename
    file_name = path_leaf(filename)
    path = filename
    imarray = readImage(path, file_name)
    print 'Process Image Complete'
    writeExcel(file_name, imarray)
    print '**** File Saved *****'

def readImage(filepath,file_name):
    image_tiff = io.imread(filepath)
    image_tiff = np.transpose(image_tiff)
    imarray = np.array(image_tiff, dtype='S32')
    return imarray

def writeExcel(filename,imarray):
    save_name = filename+'.xlsx';
    workbook = xlsxwriter.Workbook(save_name)
    worksheet = workbook.add_worksheet()
    row = 0
    for col, data in enumerate(imarray):
        worksheet.write_column(row, col, data)
    workbook.close()

def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)


if __name__ == "__main__":
    main()