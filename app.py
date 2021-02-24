# -*- coding: utf-8 -*-
"""
Created on Wed Feb 24 20:08:08 2021

@author: Mi
"""

from tkinter import *
from tkinter import filedialog
#from os.path import join, abspath
#from os import listdir
from comtypes.client import CreateObject
from tqdm import tqdm


#%%

def convertor(src, dst):
  print('src', src)
  word = CreateObject('Word.Application')
  doc = word.Documents.Open(src)
  doc.SaveAs(dst, FileFormat=17)
  doc.Close()
  word.Quit()

#%%
  
def main(files):
#  if not isinstance(files, list):
#    file = []
#    file.append(files)
#  else:
#    file = files
   
  print(files, type(files))
  file = list(filter(lambda x: True if x.endswith(".docx") or x.endswith(".doc") else False, files))
  print(file, type(file))
  if not len(file):
    print("There is no '.doc' or '.docx' files in directory")
  else:
    for src_p in tqdm(file):
      dst_p = src_p.replace('docx', 'pdf') if '.docx' in src_p else src_p.replace('doc', 'pdf')
      print(src_p, dst_p)
      convertor(src_p, dst_p)

    


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
      
      #file = list(filter(lambda x: True if x.endswith(".docx") or x.endswith(".doc") else False, self.files))
      file = ['C:/Users/Mi/Downloads/zqwe/+Абдулмуталибова АШ.docx'.replace('/', '//')]
      print(file, type(file))
      if not len(file):
        print("There is no '.doc' or '.docx' files in directory")
      else:
        for src_p in file: #(tqdm)
          dst_p = src_p.replace('docx', 'pdf') if '.docx' in src_p else src_p.replace('doc', 'pdf')
          print('here', src_p, dst_p)
          word = CreateObject('Word.Application')
          doc = word.Documents.Open(src_p)
          doc.SaveAs(dst_p, FileFormat=17)
          doc.Close()
          word.Quit()
          #convertor(src_p, dst_p)
      #main(self.files)
      self.files = None
    #else:
    #  pass
    #  # system message of error
      
  def createWidgets(self):
    
    self.QUIT = Button(self)
    self.QUIT["text"] = "QUIT"
    self.QUIT["fg"]   = "red"
    self.QUIT["command"] = self.quit
    self.QUIT.pack({"side": "left"})

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
root.geometry("200x200")
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



