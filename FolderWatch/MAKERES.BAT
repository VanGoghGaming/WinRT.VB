@ECHO OFF
CLS
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo FolderWatch.res FolderWatch.rc
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" prjFolderWatch.vbp /make /d bInIDE=0
PAUSE