# -*- coding: utf-8 -*-
"""
Created on Wed Feb 24 19:36:46 2021

@author: Mi
"""

#doc -> pdf 

from os.path import join, abspath
from os import listdir
#import sys
import comtypes.client
from tqdm import tqdm

#%%
def convertor(src, dst):
  word = comtypes.client.CreateObject('Word.Application')
  doc = word.Documents.Open(src)
  doc.SaveAs(dst, FileFormat=17)
  doc.Close()
  word.Quit()

#%%

def main(files):
  if not isinstance(files, list):
    file = []
    file.append(files)
  else:
    file = files
    
  file = list(filter(lambda x: True if x.endswith(".docx") or x.endswith(".doc") else False, listdir(file)))
  if not len(files):
    print("There is no '.doc' or '.docx' files in directory")
  else:
    for src_p in tqdm(files):
      dst_p = src_p.replace('docx', 'pdf') if '.docx' in src_p else src_p.replace('doc', 'pdf')
      convertor(src_p, dst_p)


#%%

