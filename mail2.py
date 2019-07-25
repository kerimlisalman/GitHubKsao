#pip install pypiwin32
import os
import win32com.client
Pathname=r'c://Temp//'

dir = r"C:\Temp"

file_list = os.listdir(dir)

for file in file_list:
    if file.endswith(".msg"):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem(dir + "/" + file)
        att=msg.Attachments
        for i in att:
            i.SaveAsFile(os.path.join(r'c:\Temp', i.FileName))#Saves the file with the attachment name

folder = r'c:\Temp'
for the_file in os.listdir(folder):
    file_path = os.path.join(folder, the_file)
    try:
         if the_file.endswith(".msg"):
            if os.path.isfile(file_path):
                os.unlink(file_path)
            #elif os.path.isdir(file_path): shutil.rmtree(file_path)
    except Exception as e:
        print(e)