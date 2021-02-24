# -*- coding: utf-8 -*-
"""
Created on Wed Feb 24 20:08:08 2021

@author: Mi
"""

from tkinter import *
from tkinter import filedialog
from doc_pdf_convertor import main

#%%

    


#%%

class Application(Frame):
  def browseFiles(self):
    self.files = filedialog.askopenfilenames(initialdir = "/",
                                            title = "Select a File",
                                            filetypes = (("docx", "*.docx*"),
                                                         ("doc", "*.doc*"),
                                                         ("all files", "*.*"))
                                            )
      
    # Change label contents
    #label_file_explorer.configure(text="File Opened: "+filename)
    #return filename
  
  def run(self):
    if self.files != None:
      print(self.files)
      main(self.files)
      self.files = None
    else:
      # system message of error
      
  def createWidgets(self):
    
    self.QUIT = Button(self)
    self.QUIT["text"] = "QUIT"
    self.QUIT["fg"]   = "red"
    self.QUIT["command"] = self.quit
    self.QUIT.pack({"side": "left"})

    #self.browse = Button(self)
    #Button(bottomframe, text="Black", fg="black")
    #blackbutton.pack( side = BOTTOM)
    #self.browse.pack({"side": "left"})


    self.hi_there = Button(self)
    self.hi_there["text"] = "Browse Files",
    self.hi_there["command"] = self.browseFiles
    self.hi_there.pack({"side": "left"})

    self.conv = Button(self)
    self.conv["text"] = "Convert!",
    self.conv["command"] = self.run
    self.conv.pack({"side": "top"})
    #self.conv.pack(side=BOTTOM)


  def __init__(self, master=None):
    self.files = None
    Frame.__init__(self, master)
    self.pack()
    self.createWidgets()


root = Tk()
root.title("DOC -> PDF")
root.geometry("500x200")
#root.geometry("500x500")
#Set window background color
#root.config(background = "white")
app = Application(master=root)
app.mainloop()
root.destroy()




#%%%

  
## Function for opening the 
## file explorer window
#      
#                                                                                                  
## Create the root window
#window = Tk()
#  
## Set window title
#window.title('File Explorer')
## Set window size
#window.geometry("500x500")
##Set window background color
#window.config(background = "white")
#  
## Create a File Explorer label
#label_file_explorer = Label(window, 
#                            text = "File Explorer using Tkinter",
#                            width = 100, height = 4, 
#                            fg = "blue")
#  
#      
#button_explore = Button(window, 
#                        text = "Browse Files",
#                        command = browseFiles) 
#  
#button_exit = Button(window, 
#                     text = "Exit",
#                     command = exit) 
#  
## Grid method is chosen for placing
## the widgets at respective positions 
## in a table like structure by
## specifying rows and columns
#label_file_explorer.grid(column = 1, row = 1)
#  
#button_explore.grid(column = 1, row = 2)
#  
#button_exit.grid(column = 1,row = 3)
#  
## Let the window wait for any events
#window.mainloop()



#%%


