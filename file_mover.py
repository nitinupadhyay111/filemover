
__author__ = "Nitin U."

""" file_mover.py: The purpose of this program is to read an excel spreadsheet
then from that spreadsheet it will move the files from a source folder to a 
destination folder. The source and the destination folder will be in 2 different
columns. It will also check if the destination folder exists or not. if does 
not exist then it will create that folder """

""" 
Steps
1. Read the excel file and then access the sheet
2. Create 2 lists 1 for source and 1 for destination
3. Loop through the column and create the folders if they don't exist
4. Move the files from the source location to target using shutil library 

"""

import os
import shutil
import glob
import xlrd

  
# Parent Directory path 
parent_dir = r"C:\\Users\\nitin\\Dropbox (Personal)\\Swinburne Statistics\\STA70005-Survey Sampling\\2020\\files\\pdf - Copy\\"

# Directory to be created under the parent directory

def create_directory(par_dir, dir):
    path = os.path.join(par_dir, dir) 
    os.mkdir(path) 
    

def check_if_exists(par_dir, dir):

    """ This function will check whether the specified path is an existing directory or not this will return a list """
      
    path = os.path.join(par_dir, dir)
    
    isdir = os.path.isdir(path)  
    
    return isdir
    
def get_file_names(folder_name):
    
    """ This will get the full path of all the files in the folder """
    
    file_path_list = glob.glob(folder_name)
    
    return file_path_list    


def main():
    
    file_to_read_from = r"C:\Users\nitin\Dropbox (Personal)\Swinburne Statistics\STA70005-Survey Sampling\2020\Output.xlsx"
    
    book = xlrd.open_workbook(file_to_read_from)
    
    sheet = book.sheet_by_name("Sheet1")
  
    for i in range(1, len(sheet.col(0))):

        src_file = sheet.row(i)[0].value # column 0: this column contains all the original path of the files.
        dest_folder = sheet.row(i)[7].value # column 7: this column contains the full path of the destination folder NOT THE FILES, only Folders
        folder = sheet.row(i)[6].value # column 6: this column contains the name of the destination folder

        try:
            if check_if_exists(parent_dir, folder) == True:

                print("Directory '% s' already exists under the parent directory chutiye" % folder)

            else:

                create_directory(parent_dir, folder)
                print("===========================")   
                print("Directory '% s' created under the parent directory" % folder)   
                print("===========================")   

            shutil.move(r""+src_file, r""+str(dest_folder))

            print("file % s moved" % src_file)

        except Exception as why:
            
            print(why)

   
if __name__ == "__main__":
    main()
  