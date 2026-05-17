VERSION 5.00
Begin VB.Form frmMediaPlayer 
   Caption         =   "Media Player - Double click to select a media file for playing"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMediaPlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Media Player"
End
Attribute VB_Name = "frmMediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDoubleTappedEventHandler

Private m_hWndIsland As LongPtr, m_DesktopWindowXamlSource As IDesktopWindowXamlSource, m_MediaPlayerElement As IMediaPlayerElement

Private Sub Form_Load()
Dim DesktopWindowXamlSourceNative As IDesktopWindowXamlSourceNative, UIElement As IUIElement
    Set m_DesktopWindowXamlSource = NewObject("DesktopWindowXamlSource"): Set DesktopWindowXamlSourceNative = m_DesktopWindowXamlSource
    With DesktopWindowXamlSourceNative
        If .AttachToWindow(hWnd) = S_OK Then
            m_hWndIsland = .WindowHandle
            Set m_MediaPlayerElement = NewObject("MediaPlayerElement"): Set UIElement = m_MediaPlayerElement: UIElement.AddDoubleTapped Me
            m_MediaPlayerElement.AutoPlay = APITRUE: m_MediaPlayerElement.AreTransportControlsEnabled = APITRUE
            Set m_DesktopWindowXamlSource.Content = m_MediaPlayerElement
        End If
    End With
End Sub

Private Sub Form_Resize()
    SetWindowPos m_hWndIsland, HWND_TOP, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
End Sub

Private Sub IDoubleTappedEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As IDoubleTappedRoutedEventArgs)
Dim StorageItem As IStorageItem
    With New cFilePicker
        If .SetOwnerWindow(hWnd) Then
            .SetFileTypeFilters ".mp4", ".mkv", ".wmv", ".avi", ".mov", ".3gp", ".mp3", ".m4a", ".wma", ".flac", ".wav"
            Select Case .ShowFileOpenPicker(StorageItem)
                Case FilePickerResult_OK
                    Caption = Tag & " - " & .FileName: m_MediaPlayerElement.SetMediaPlayer Nothing
                    Set m_MediaPlayerElement.Source = MediaSourceStatics.CreateFromStorageFile(StorageItem)
                Case FilePickerResult_Canceled: MsgBox "Canceled by user!", vbOKOnly + vbExclamation, App.Title
                Case FilePickerResult_Timeout: MsgBox "Timeout expired!", vbOKOnly + vbExclamation, App.Title
                Case FilePickerResult_Error: MsgBox "FilePicker can't run elevated (as administrator)!", vbOKOnly + vbCritical, App.Title
            End Select
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_MediaPlayerElement.SetMediaPlayer Nothing: Set m_MediaPlayerElement = Nothing: Dispose m_DesktopWindowXamlSource
End Sub
