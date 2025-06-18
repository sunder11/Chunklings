If you add the code to a text file with a .bat extension and double click it in windows it will search through all the subfolders in your myfiles directory and copy all the files with a pdf extesion to a folder you specify. Dup.licate file names will be appended with _1, _2 etc so no files with the same name are overwritten.

Uses: to put all your pdf's in one folder so docling can chunk your data 

To use edit the .bat file and change the path to source directory and target directory

rem Set your folders:
set "source_folder=C:\AI\Copypdf"
set "target_folder=C:\AI\2"
