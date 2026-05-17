VERSION 5.00
Begin VB.Form frmMediaPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Player"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3855
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
   MaxButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSlower 
      Caption         =   "-"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Play slower"
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdFaster 
      Caption         =   "+"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      ToolTipText     =   "Play faster"
      Top             =   720
      Width           =   255
   End
   Begin VB.HScrollBar hscVolume 
      Height          =   495
      LargeChange     =   10
      Left            =   1080
      Max             =   100
      TabIndex        =   2
      Top             =   120
      Value           =   50
      Width           =   2655
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Pause"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblPlaybackRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rate: "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Tag             =   "Rate: "
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblDuration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Duration: "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Tag             =   "Duration: "
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmMediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements cMediaPlayerEvents

Private MediaPlayer As IMediaPlayer

Private Sub cmdFaster_Click()
    With MediaPlayer
        Select Case .PlaybackRate + 0.25
            Case Is > 2
            Case Else: .PlaybackRate = .PlaybackRate + 0.25
        End Select
    End With
End Sub

Private Sub cmdPause_Click()
    MediaPlayer.Pause
End Sub

Private Sub cmdPlay_Click()
    MediaPlayer.Play
End Sub

Private Sub cmdSlower_Click()
    With MediaPlayer
        Select Case .PlaybackRate - 0.25
            Case Is < 0.25
            Case Else: .PlaybackRate = .PlaybackRate - 0.25
        End Select
    End With
End Sub

Private Sub cMediaPlayerEvents_EventHandler(ByVal EventIdentifier As MEDIAPLAYER_EVENT_IDENTIFIERS, ByVal Sender As IInspectable, Optional ByVal IInspectableEventArgs As IInspectable, Optional ByVal MediaPlayerFailedEventArgs As IMediaPlayerFailedEventArgs, Optional ByVal MediaPlayerRateChangedEventArgs As IMediaPlayerRateChangedEventArgs)
Dim MediaPlayer3 As IMediaPlayer3, MediaPlaybackSession As IMediaPlaybackSession, lMinutes As Long, lSeconds As Long, lMilliseconds As Long
    Select Case EventIdentifier
        Case MediaPlaybackSessionEvent_PositionChanged
            Set MediaPlaybackSession = Sender
            With MediaPlaybackSession
                lMinutes = (.NaturalDuration - .Position) \ 60000: lSeconds = (.NaturalDuration - .Position) \ 1000 Mod 60: lMilliseconds = (.NaturalDuration - .Position) Mod 1000
                lblDuration = lblDuration.Tag & Format$(lMinutes, "00") & ":" & Format$(lSeconds, "00") & "." & Format$(lMilliseconds, "000")
            End With
        Case MediaPlayerEvent_CurrentStateChanged
            Set MediaPlayer3 = Sender
            Select Case MediaPlayer3.PlaybackSession.PlaybackState
                Case MediaPlaybackState_Paused: cmdPause.Enabled = False: cmdPlay.Enabled = True: cmdPlay.SetFocus
                Case MediaPlaybackState_Playing: cmdPause.Enabled = True: cmdPlay.Enabled = False: cmdPause.SetFocus
            End Select
        Case MediaPlayerEvent_Ended
            Set MediaPlayer3 = Sender
            MediaPlayer3.PlaybackSession.MediaPlayer.Play
        Case MediaPlayerEvent_Failed
            With MediaPlayerFailedEventArgs
                Debug.Print .Error, GetStr(.ErrorMessage), Hex$(.ExtendedErrorCode), GetErrorMessage(.ExtendedErrorCode)
            End With
        Case MediaPlayerEvent_Opened
            Set MediaPlayer3 = Sender
            With MediaPlayer3.PlaybackSession
                lMinutes = .NaturalDuration \ 60000: lSeconds = .NaturalDuration \ 1000 Mod 60: lMilliseconds = .NaturalDuration Mod 1000
                lblDuration = lblDuration.Tag & Format$(lMinutes, "00") & ":" & Format$(lSeconds, "00") & "." & Format$(lMilliseconds, "000")
                .MediaPlayer.Play
            End With
        Case MediaPlayerEvent_RateChanged
            lblPlaybackRate = lblPlaybackRate.Tag & MediaPlayerRateChangedEventArgs.NewRate * 100 & "%"
    End Select
End Sub

Private Sub Form_Load()
Dim sFile As String, MediaPlayer3 As IMediaPlayer3
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    sFile = "https://github.com/prof3ssorSt3v3/media-sample-files/raw/refs/heads/master/jimmy-coffee.mp3" ' file from the Web
    Set MediaPlayer = NewObject("MediaPlayer"): Set MediaPlayer3 = MediaPlayer: hscVolume_Change
    MediaPlayer3.PlaybackSession.AddPositionChanged NewEventHandler(MediaPlaybackSessionEvent_PositionChanged)
    With MediaPlayer
        .AddCurrentStateChanged NewEventHandler(MediaPlayerEvent_CurrentStateChanged)
        .AddMediaEnded NewEventHandler(MediaPlayerEvent_Ended)
        .AddMediaFailed NewEventHandler(MediaPlayerEvent_Failed)
        .AddMediaOpened NewEventHandler(MediaPlayerEvent_Opened)
        .AddMediaPlayerRateChanged NewEventHandler(MediaPlayerEvent_RateChanged)
        .SetUriSource UriFactory.CreateUri(StrRef(sFile))
        lblPlaybackRate = lblPlaybackRate.Tag & .PlaybackRate * 100 & "%"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dispose MediaPlayer
End Sub

Private Sub hscVolume_Change()
    MediaPlayer.Volume = hscVolume / 100
End Sub

Private Function NewEventHandler(ByVal EventIdentifier As MEDIAPLAYER_EVENT_IDENTIFIERS) As cMediaPlayerEvents
    With New cMediaPlayerEvents
        Set NewEventHandler = .This(Me, EventIdentifier)
    End With
End Function
