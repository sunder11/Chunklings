Use These Windows Utilities to convert your emails from your outlook .pst file to html

Source code and credit to the following:

https://github.com/Dijji/XstReader
--------------------------------------------------------------------

Export emails from reader:
XstReader.exe (double click /open .pst file/ right click on folder want to export/export emails/choose export folder  

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