If you add the code to a text file with a .bat extension and double click it in windows with 64bit LibreOffice installed and place it in your myfiles folder (that of course has a bunch of subfolders in it) it will go through all the subfolders and convert .wpd to .docx and place all the files in a folder C:\AI\1 (modify the target folder if you like). 

The files with the same name will be appended with a _1 _2 etc. 

Thanks to Open AI gpt-4.5-preview-2025-02-27 for helping me do it. I think this code will be useful to implement RAG and to convert files in a format that docling (IBM's opensource project) can use to chunk your data for your vector database. Probably easier to parse your files if they are all in the same folder when I working on AI Agents.