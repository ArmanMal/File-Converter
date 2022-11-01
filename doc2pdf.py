from docx2pdf import convert
import os, webbrowser
from tkinter import Tk, filedialog

#Defining paths
path = "C:\\..."
downloads_path = "C:\\..."

#Selecting word files to convert
downloadfiles = os.listdir(downloads_path)
root =Tk()
root.withdraw()
root.attributes('-topmost', True)
word_files = filedialog.askopenfilenames(initialdir=downloads_path, filetypes=[("Word files", "*doc;*.docx")])
print("Files to convert:")

for w in word_files:
    basename = os.path.basename(w)
    print(basename)
    file_to_move = downloads_path + basename
    new_path = path + basename
    os.rename(file_to_move, new_path)
   
#convert PDFs
convert(path)

#Open PDFs/change destination and delete word files
os.chdir(path)
folder = os.listdir()

print("PDFs downloaded:")
for f in folder:
    if f.endswith(".pdf"):
        print(f)
        file_to_move = path + f
        new_path = downloads_path + f

        os.rename(file_to_move, new_path)
        webbrowser.open(new_path)

    else:
        os.remove(f)
