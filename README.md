# String_Finder_File_List_Generator
Python3 application to:  (1) search for specified string in files in a specified directory; (2) generate .txt file of all files in a specified directory

#### String search
Specify the string you want to search for in the specified directory and choose the file types you want to search on, and it will output a list of files which contain the string.
Sometimes the program can't read in files to perform the search for whatever reason (most common scenario I found was with pdfs where there was some kind of internal encryption in the file).  In these situations, the program will skip those files, but let you know which ones it skipped so that you can check those manually.

#### File list generator
Specify the directory you want to list all the files contained in them, and an output file will be generated which lists:
- all the files with absolute filepaths
- all the files using just the filenames

Note that if there are any subdirectories inside the specified directory, all files inside of subdirectories will included in the string search feature or the file list generation feature.

#### Libraries used
- tkinter
- PyPDF2
- docx


## Author
Peter Chung
