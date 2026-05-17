VERSION 5.00
Begin VB.Form frmStreamSockets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stream Socket SSL Demo"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStreamSockets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   8415
   End
   Begin VB.TextBox txtMessages 
      Height          =   6855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   8415
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "587"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtHost 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "smtp.gmail.com"
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmStreamSockets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cSocketClient As cSocketClient, bBase64Encode As Boolean
Attribute cSocketClient.VB_VarHelpID = -1

Private Sub cmdConnect_Click()
    If cSocketClient.Connect(txtHost, txtPort) Then cmdConnect.Enabled = False: txtHost.Enabled = False: txtPort.Enabled = False: txtMessage.Enabled = True
End Sub

Private Sub cSocketClient_DataReceived(sData As String)
    If InStr(sData, "220 2.0.0 Ready to start TLS") > 0 Then ' In this example the connection is upgraded to TLS after sending the STARTTLS command
        cSocketClient.UpgradeToSsl
    ElseIf InStr(sData, "334 ") > 0 Then ' Usernames and Passwords are Base64 encoded
        sData = "334 " & Base64Decode(Mid$(Left$(sData, Len(sData) - 2), 5)): bBase64Encode = True
    ElseIf InStr(sData, "535-5.7.8") > 0 Or InStr(UCase$(sData), "QUIT") > 0 Then
        bBase64Encode = False
    End If
    txtMessages = txtMessages & sData & vbNewLine: txtMessage.SetFocus
End Sub

Private Sub Form_Load()
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    cPop.ParentWindow = hWnd
    Set cSocketClient = New cSocketClient
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
Dim sMessage As String
    If KeyAscii = vbKeyReturn Then
        sMessage = txtMessage
        If Len(sMessage) Then If cSocketClient.SendMessage(IIf(bBase64Encode, Base64Encode(sMessage), sMessage) & vbNewLine) Then txtMessage = vbNullString: cSocketClient_DataReceived sMessage
        KeyAscii = 0
    End If
End Sub

Private Function Base64Decode(sData As String) As String
    With CryptographicBufferStatics
        Base64Decode = GetStr(.ConvertBinaryToString(UnicodeEncoding_UTF8, .DecodeFromBase64String(StrRef(sData))))
    End With
End Function

Private Function Base64Encode(sData As String) As String
    With CryptographicBufferStatics
        Base64Encode = GetStr(.EncodeToBase64String(.ConvertStringToBinary(StrRef(sData), UnicodeEncoding_UTF8)))
    End With
End Function
