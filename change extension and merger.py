import os
import glob
import datetime

"""
Change extension works only if the file has only text file and py file
"""
def add_txt():
    extension = ".txt"
    for filename in os.listdir():
        if ".py" not in filename and ".xlsx" not in filename: 
            os.rename(filename, filename + extension)
    
#If done by mistake:
def fix():
    extension = ".txt"
    for file in os.listdir():
        os.rename(file, file.replace(extension, ""))

today = datetime.date.today().strftime('%Y-%m-%d')

#MERGE TXT FILES--------------------------------------------------------------
def merge_files():
    #Read all text files and combine all files into a text file
    all_files = glob.glob("*.txt")
    
    #Create a merging txt file 
    batch_name = input("What is the name of the batch? ")
    merged_file = batch_name + " merged on " + today + ".txt"
    with open(merged_file, "w") as outfile:
        for file in all_files:
            with open(file, "r", encoding = "ISO-8859-1") as infile:
                outfile.write(infile.read())

print("Delete the merged text file created before. Make sure every file \
needs to be merged. We don't want replicate data.")
    
ans = input("Do you want to merge all of the text files?[Y/N] ")
ans_list = ["Y", "y", "N", "n"]

while ans not in ans_list:
    print("Please only put Y or N.")
    ans = input("Do you want to merge all of the text files?[Y/N] ")

if ans == "Y" or ans == "y":
    merge_files()
else:
    pass

#User converts the merged text file into an excel file---------------------
print("Please convert the resulting text file into an excel file. Keep the \
same file name and put the excel file in the same folder if you want to extract\
efficiency.")



    
    
    