VERSION 5.00
Begin VB.Form frmFolderWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FolderWatch Demo"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17160
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFolderWatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   17160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFolderWatch 
      Height          =   8460
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   16935
   End
End
Attribute VB_Name = "frmFolderWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents FileQuery As cFileQuery
Attribute FileQuery.VB_VarHelpID = -1

Private Sub FileQuery_FileChanged(ByVal sPath As String, ByVal FileChange As FileChanges, ByVal FileSize As UInt64, ByVal OldSize As UInt64, ByVal DateModified As Date, ByVal OldDate As Date, ByVal FileAttributes As FileAttributes, ByVal OldAttributes As FileAttributes)
    txtFolderWatch = txtFolderWatch & "FileChanged: " & sPath & " | " & FileChange & " | " & FileSize & " | " & OldSize & " | " & DateModified & " | " & OldDate & " | " & FileAttributes & " | " & OldAttributes & vbNewLine & vbNewLine
End Sub

Private Sub FileQuery_FileCreated(ByVal sPath As String)
    txtFolderWatch = txtFolderWatch & "FileCreated: " & sPath & vbNewLine & vbNewLine
End Sub

Private Sub FileQuery_FileDeleted(ByVal sPath As String)
    txtFolderWatch = txtFolderWatch & "FileDeleted: " & sPath & vbNewLine & vbNewLine
End Sub

Private Sub Form_Load()
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    With New cFileQuery
        Set FileQuery = .This(AppPath & "\WatchedFolder")
    End With
End Sub
