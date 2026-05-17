VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventsSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IExceptionRoutedEventHandler
Implements IEventHandlerIInspectable
Implements IPointerEventHandler
Implements IRoutedEventHandler
Implements ITappedEventHandler

Public Enum EVENT_IDENTIFIERS
    Event_None
'    IInspectableEvents:
    Event_ClipboardContentChanged
    Event_FlyoutClosed
    Event_FlyoutOpening
'    PointerRoutedEvents:
    Event_PointerCanceled
    Event_PointerCaptureLost
    Event_PointerEntered
    Event_PointerExited
    Event_PointerMoved
    Event_PointerPressed
    Event_PointerReleased
    Event_PointerWheelChanged
'    RoutedEvents:
    Event_Click
    Event_Loaded
    Event_ImageOpened
'    ExceptionRoutedEvent:
    Event_ImageFailed
'    TappedRoutedEvent:
    Event_Tapped
End Enum

Private m_EventIdentifier As EVENT_IDENTIFIERS, m_Callback As cEventsSink

Friend Function This(Callback As cEventsSink, EventIdentifier As EVENT_IDENTIFIERS) As cEventsSink
    Set m_Callback = Callback: m_EventIdentifier = EventIdentifier: Set This = Me
End Function

Public Sub EventHandler(ByVal EventIdentifier As EVENT_IDENTIFIERS, ByVal Sender As IInspectable, Optional ByVal PointerRoutedEventArgs As IPointerRoutedEventArgs, Optional ByVal RoutedEventArgs As IRoutedEventArgs, Optional ByVal ExceptionRoutedEventArgs As IExceptionRoutedEventArgs, Optional ByVal TappedRoutedEventArgs As ITappedRoutedEventArgs)

End Sub

Private Sub IEventHandlerIInspectable_Invoke(ByVal Sender As IInspectable, ByVal Args As IInspectable)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender
End Sub

Private Sub IExceptionRoutedEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As IExceptionRoutedEventArgs)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, ExceptionRoutedEventArgs:=Args
End Sub

Private Sub IPointerEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As IPointerRoutedEventArgs)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, PointerRoutedEventArgs:=Args
End Sub

Private Sub IRoutedEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As IRoutedEventArgs)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, RoutedEventArgs:=Args
End Sub

Private Sub ITappedEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As ITappedRoutedEventArgs)
    If Not m_Callback Is Nothing Then m_Callback.EventHandler m_EventIdentifier, Sender, TappedRoutedEventArgs:=Args
End Sub
