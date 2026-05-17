VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IAsyncOperationCompletedHandlerIUICommand

Private m_MessageDialog As IMessageDialog, m_PopupMenu As IPopupMenu, m_Callback As IWeakRef, m_HandlerId As Variant, m_hWndParent As LongPtr, m_Position As POINTF, m_Await As New cAwait

Friend Property Get Callback() As cPop
    If Not m_Callback Is Nothing Then If m_Callback.StillAlive Then Set Callback = m_Callback.Resolve
End Property

Friend Property Set Callback(Callback As IWeakRefSrc)
    If Not Callback Is Nothing Then Set m_Callback = Callback.GetWeakRef
End Property

Friend Property Get HandlerId() As Variant
    HandlerId = m_HandlerId
End Property

Friend Property Let HandlerId(HandlerId As Variant)
    m_HandlerId = HandlerId
End Property

Friend Function MessageBox(sMessage As String, sTitle As String, ParamArray vaButtons() As Variant) As Variant
Dim i As Long, UICommand As IUICommand, ButtonId As Long
    If Len(sTitle) Then
        Set m_MessageDialog = MessageDialogFactory.CreateWithTitle(StrRef(sMessage), StrRef(sTitle))
    Else
        Set m_MessageDialog = MessageDialogFactory.Create(StrRef(sMessage))
    End If
    If InitializeWithWindow(m_MessageDialog, m_hWndParent) Then
        With m_MessageDialog
            With .Commands
                For i = LBound(vaButtons) To UBound(vaButtons)
                    ButtonId = Val(vaButtons(i))
                    Select Case vaButtons(i)
                        Case vbOK: vaButtons(i) = "&Okay"
                        Case vbCancel: vaButtons(i) = "&Cancel"
                        Case vbYes: vaButtons(i) = "&Yes"
                        Case vbNo: vaButtons(i) = "&No"
                        Case vbAbort: vaButtons(i) = "&Abort"
                        Case vbRetry: vaButtons(i) = "&Retry"
                        Case vbIgnore: vaButtons(i) = "&Ignore"
                        Case vbContinue: vaButtons(i) = "Co&ntinue"
                        Case vbTryAgain: vaButtons(i) = "&Try Again"
                        Case Else: ButtonId = i: vaButtons(i) = CStr(vaButtons(i))
                    End Select
                    If .Append(UICommandFactory.Create(StrRef(vaButtons(i)))) = S_OK Then Set .GetAt(i).Id = Box(ButtonId) Else Exit For
                Next i
            End With
            .DefaultCommandIndex = 0: .CancelCommandIndex = i - 1
            If m_Callback Is Nothing Then
                If m_Await.Await(.ShowAsync) = AsyncStatus_Completed Then Set UICommand = m_Await.GetResults: MessageBox = Unbox(UICommand.Id)
            Else
                Set .ShowAsync.Completed = Me
            End If
        End With
    End If
End Function

Friend Property Get ParentWindow() As LongPtr
    ParentWindow = m_hWndParent
End Property

Friend Property Let ParentWindow(hWndParent As LongPtr)
    m_hWndParent = hWndParent
End Property

Friend Function PopupMenu(ParamArray vaMenuOptions() As Variant) As Long
Dim i As Long, UICommand As IUICommand
    PopupMenu = -1: Set m_PopupMenu = NewObject("PopupMenu")
    If InitializeWithWindow(m_PopupMenu, m_hWndParent) Then
        With m_PopupMenu
            With .Commands
                For i = LBound(vaMenuOptions) To UBound(vaMenuOptions)
                    If .Append(UICommandFactory.Create(StrRef(vaMenuOptions(i)))) = S_OK Then Set .GetAt(i).Id = Box(i) Else Exit For
                Next i
            End With
            If m_Callback Is Nothing Then
                If m_Await.Await(.ShowAsync(m_Position.X, m_Position.Y)) = AsyncStatus_Completed Then Set UICommand = m_Await.GetResults: If Not UICommand Is Nothing Then PopupMenu = Unbox(UICommand.Id)
            Else
                #If Win64 Then
                    Dim Position As PointByVal: vbaCopyBytes LenB(Position), ByVal VarPtr(Position), ByVal VarPtr(m_Position)
                    Set .ShowAsync(Position).Completed = Me
                #Else
                    Set .ShowAsync(m_Position.X, m_Position.Y).Completed = Me
                #End If
            End If
        End With
    End If
End Function

Friend Sub SetPopupMenuPosition(ByVal PositionX As Single, ByVal PositionY As Single)
    m_Position.X = PositionX: m_Position.Y = PositionY
End Sub

Friend Function This(Optional hWndParent As LongPtr, Optional Callback As IWeakRefSrc, Optional HandlerId As Variant) As cPop
    If hWndParent <> vbNullPtr Then m_hWndParent = hWndParent
    If Not Callback Is Nothing Then Set m_Callback = Callback.GetWeakRef
    If Not IsMissing(HandlerId) Then m_HandlerId = HandlerId
    Set This = Me
End Function

Public Sub PopResult(HandlerId As Variant, ResultId As Variant, ResultLabel As Variant)

End Sub

Private Sub IAsyncOperationCompletedHandlerIUICommand_Invoke(ByVal AsyncOperation As IAsyncOperationIUICommand, ByVal AsyncStatus As AsyncStatus)
Dim Callback As cPop, UICommand As IUICommand
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            Set UICommand = AsyncOperation.GetResults
            If Not m_Callback Is Nothing And Not UICommand Is Nothing Then
                If m_Callback.StillAlive Then
                    Set Callback = m_Callback.Resolve
                    With UICommand: Callback.PopResult m_HandlerId, Unbox(.Id), GetStr(.Label): End With
                End If
            End If
    End Select
End Sub
