@ECHO OFF
CLS
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo ShareUI.res ShareUI.rc
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" prjShareUI.vbp /make /d bInIDE=0
PAUSE