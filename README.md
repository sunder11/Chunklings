If you are working on AI Agents and are trying to get your data ready for chunking with docling or the like, the following utilities may be helpful.

**ConvertWPDtoDocx:**
Convert your Wordperfect documents to .docx:

Doubleclick covert.bat in windows after installing 64bit LibreOffice and after placing the batch file in your myfiles folder (that of course has a bunch of subfolders in it). It will go through all the subfolders and convert .wpd to .docx and place all the files in a folder C:\AI\1 (modify the target folder if you like). 

The files with the same name will be appended with a _1 _2 etc. 

I think this code will be useful to implement RAG and to convert files in a format that docling (IBM's opensource project) can use to chunk your data for your vector database. Probably easier to parse your files if they are all in the same folder.

**CopyPDF:** Copy all your PDF files to a single directory.

Doubleclick copy_pdf.bat in windows after placing the batch file in your myfiles folder it will search through all the subfolders and copy all the files with a .pdf extesion to a folder you specify. Duplicate file names will be appended with _1, _2 etc so no files with the same name are overwritten.

To Use: Edit in a text editor like notepad and change the path to source directory and target directory.

rem Set your folders:<br>
set "source_folder=C:\AI\Copypdf"<br>
set "target_folder=C:\AI\2"

Thanks to OpenAI gpt-4.5-preview-2025-02-27 for helping me with the 2 utilities above (or I should say for writing the code for about $5.00).<br>

**ExcelToJson** Convert your Excel .xlsx to .Json format.<br> 

**CRMbat**  Convert your CRM to html file. Modify the fields (4 places) to match yours (and their position number), and at the output section where you will see !FIELD_NAME!.  It will writeout all of your database records numbering them in an html file with a sentence describing the record.I found it works better to export your CRM in a format Excel can read and save it as an .xlsx file, but this cound be useful.<br>

**SaveOutlookEmailsToDocx:**  Copy and Paste this Visual Basic code and save it as a macro in outlook and add it to the Outlook Ribbon (button on the toolbar). To Use: Modify the path to where you want to save the .docx file. Select the emails you want to copy, click the button, it will prompt for a file_name, then it will copy all the emails selected to one file called "file_name.docx". 



