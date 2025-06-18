If you are working on AI Agents and are trying to get your data ready for chunking with docling or the like, the following utilities may be helpful to someone I am thinking.

**ConvertWPDtoDocx:**
Convert your Wordperfect documents to .docx:

Doubleclick covert.bat in windows after installing 64bit LibreOffice and after placing the batch file in your myfiles folder (that of course has a bunch of subfolders in it). It will go through all the subfolders and convert .wpd to .docx and place all the files in a folder C:\AI\1 (modify the target folder if you like). 

The files with the same name will be appended with a _1 _2 etc. 

Thanks to OpenAI gpt-4.5-preview-2025-02-27 for helping me do it. I think this code will be useful to implement RAG and to convert files in a format that docling (IBM's opensource project) can use to chunk your data for your vector database. Probably easier to parse your files if they are all in the same folder.

**CopyPDF:** Copy all your PDF files to a single directory.

Doubleclick copy_pdf.bat in windows after placing the batch file in your myfiles folder it will search through all the subfolders and copy all the files with a .pdf extesion to a folder you specify. Duplicate file names will be appended with _1, _2 etc so no files with the same name are overwritten.

To Use: Edit in a text editor like notepad and change the path to source directory and target directory.

rem Set your folders:<br>
set "source_folder=C:\AI\Copypdf"<br>
set "target_folder=C:\AI\2"

**XportEmailsFromPst:**  Use These Windows Utilities to extract your emails from your outlook@example.com.pst file and convert them to html.

All the credit to:

https://github.com/Dijji/XstReader
--------------------------------------------------------------------

Export emails from reader:
XstReader.exe (double click /open .pst file/ right click on folder you  want to export/export emails/choose export folder  

-----------------------------------------------------------------------

XstExport (from commandline - modify switches in shortcut)

A command line tool for exporting emails, attachments or properties from an Outlook file.

By default, all folders in the Outlook file are exported into a directory structure that mirrors the Outlook folder structure. Options are available to specify the starting Outlook folder and to collapse all output into a single directory.

The differences from the export capabilities of the UI are: the ability to export from a subtree of Outlook folders; and the ability to export attachments only, without the body of the email.

In addition to XstExport, XstPortableExport is also provided, which is a portable version based on .Net Core 2.1 that can be run on Windows, Mac, Linux etc.

Both versions support the following options:

XstExport.exe {-e|-p|-a|-h} [-f=<Outlook folder>] [-o] [-s] [-t=<target directory>] <Outlook file name>

XstExport.exe -e -f=Sent -t=C:\AI\2 steve@mrsumppump.com.pst


Where:

-e, --email
Export in native body format (.html, .rtf, .txt) with attachments in associated folder
-- OR --
-p, --properties
Export properties only (in CSV file)
-- OR --
-a, --attachments
Export attachments only (Latest date wins in case of name conflict)
-- OR --
-h, --help
Display this help

-f=<Outlook folder>, -folder=<Outlook folder>
Folder within the Outlook file from which to export. This may be a partial path, for example "Week1\Sent"

-o, --only
If set, do not export from subfolders of the nominated folder.

-s, --subfolders
If set, Outlook subfolder structure is preserved. Otherwise, all output goes to a single directory

-t=<target directory name>, --target=<target directory name>
The directory to which output is written. This may be an absolute path or one relative to the location of the Outlook file. By default, output is written to a directory <Outlook file name>.Export.<Command> created in the same directory as the Outlook file

<Outlook file name>
The full name of the .pst or .ost file from which to export

To run the portable version, open a command line and run:

dotnet XstPortableExport.dll <options as above>

