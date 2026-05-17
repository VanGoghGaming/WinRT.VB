@ECHO OFF
CLS
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo PDFViewer.res PDFViewer.rc
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" prjPDFViewer.vbp /make /d bInIDE=0
PAUSE