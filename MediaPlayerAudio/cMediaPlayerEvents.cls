VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMediaPlayerEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITypedEventHandlerMediaPlayerIInspectable
Implements ITypedEventHandlerMediaPlayerMediaFailed
Implements ITypedEventHandlerMediaPlaybackSessionIInspectable
Implements ITypedEventHandlerMediaPlayerRateChanged

Public Enum MEDIAPLAYER_EVENT_IDENTIFIERS ' Add more events as needed
    MediaPlayerEvent_None
    MediaPlaybackSessionEvent_PositionChanged
    MediaPlayerEvent_CurrentStateChanged
    MediaPlayerEvent_Ended
    MediaPlayerEvent_Failed
    MediaPlayerEvent_Opened
    MediaPlayerEvent_RateChanged
End Enum

Private m_EventIdentifier As MEDIAPLAYER_EVENT_IDENTIFIERS, m_Callback As cMediaPlayerEvents

Friend Function This(Callback As cMediaPlayerEvents, ByVal EventIdentifier As MEDIAPLAYER_EVENT_IDENTIFIERS) As cMediaPlayerEvents
    Set m_Callback = Callback: m_EventIdentifier = EventIdentifier: Set This = Me
End Function

Public Sub EventHandler(ByVal EventIdentifier As MEDIAPLAYER_EVENT_IDENTIFIERS, ByVal Sender As IInspectable, Optional ByVal IInspectableEventArgs As IInspectable, Optional ByVal MediaPlayerFailedEventArgs As IMediaPlayerFailedEventArgs, Optional ByVal MediaPlayerRateChangedEventArgs As IMediaPlayerRateChangedEventArgs)

End Sub

Private Sub ITypedEventHandlerMediaPlaybackSessionIInspectable_Invoke(ByVal Sender As IMediaPlaybackSession, ByVal Args As IInspectable)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, IInspectableEventArgs:=Args
End Sub

Private Sub ITypedEventHandlerMediaPlayerIInspectable_Invoke(ByVal Sender As IMediaPlayer, ByVal Args As IInspectable)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, IInspectableEventArgs:=Args
End Sub

Private Sub ITypedEventHandlerMediaPlayerMediaFailed_Invoke(ByVal Sender As IMediaPlayer, ByVal Args As IMediaPlayerFailedEventArgs)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, MediaPlayerFailedEventArgs:=Args
End Sub

Private Sub ITypedEventHandlerMediaPlayerRateChanged_Invoke(ByVal Sender As IMediaPlayer, ByVal Args As IMediaPlayerRateChangedEventArgs)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, MediaPlayerRateChangedEventArgs:=Args
End Sub
