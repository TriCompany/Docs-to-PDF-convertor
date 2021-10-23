import os 
import pythoncom
 
# import comtypes.client as client 
import win32com 
import win32com.client as client 
from shutil import copyfile, rmtree 
import threading 


wdFormatPDF = 17 
files = ['MATH 137/Notes.docx', 'MATH 135/Notes.docx'] 
pythoncom.CoInitialize() 
word = client.Dispatch('Word.Application') # word = comtypes.client.CreateObject('Word.Application') 


def convert_to_pdf(in_file, word_id): 
 pythoncom.CoInitialize() 
 word = client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(word_id, pythoncom.IID_IDispatch)) 
 word.Visible = False 
 in_file = os.path.abspath(in_file) 
 out_file = in_file.replace('.docx','.pdf').replace('.doc', '.pdf') 
 doc = word.Documents.Open(in_file) 
 doc.SaveAs(out_file, FileFormat=wdFormatPDF) 
 doc.Close()


threads = [] 
for file in files: 
 word_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, word) 
 t = threading.Thread(target=convert_to_pdf, args=(file, word_id)) 
 t.start() 
 threads.append(t)

print("Creating PDFs") 
for t in threads: t.join() 
word.Quit() 

print("PDFs Created")
