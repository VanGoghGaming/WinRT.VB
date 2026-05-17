VERSION 5.00
Begin VB.Form frmWindowsHello 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Hello"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWindowsHello.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWindowsHello 
      Caption         =   "Windows Hello"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmWindowsHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cUserConsentVerifier As cUserConsentVerifier
Attribute cUserConsentVerifier.VB_VarHelpID = -1

Private Sub Form_Load()
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    Set cUserConsentVerifier = New cUserConsentVerifier: cUserConsentVerifier.ParentWindow = hWnd
End Sub

Private Sub cmdWindowsHello_Click()
    cUserConsentVerifier.ShowWindowsHello "Windows Hello says please verify your identity!"
End Sub

Private Sub cUserConsentVerifier_UserConsentVerificationResult(UserConsentVerificationResult As UserConsentVerificationResult)
    MsgBox Choose(UserConsentVerificationResult + 1, "Verified successfully!", "DeviceNotPresent", "NotConfiguredForUser", "DisabledByPolicy", "DeviceBusy", "RetriesExhausted", "Canceled by user!"), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub cUserConsentVerifier_CredentialPickerResult(UserConsentVerificationResult As UserConsentVerificationResult)
    Select Case UserConsentVerificationResult
        Case UserConsentVerificationResult_Verified: MsgBox "Verified successfully!", vbOKOnly + vbInformation, App.Title
        Case UserConsentVerificationResult_Canceled: MsgBox "Canceled by user!", vbOKOnly + vbExclamation, App.Title
        Case Else: MsgBox "Incorrect credentials!", vbOKOnly + vbCritical, App.Title
    End Select
End Sub
