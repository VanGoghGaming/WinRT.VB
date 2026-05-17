@ECHO OFF
CLS
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo GIFPlayer.res GIFPlayer.rc
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" prjGIFPlayer.vbp /make /d bInIDE=0
PAUSE