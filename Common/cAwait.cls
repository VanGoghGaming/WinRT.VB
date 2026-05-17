VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAwait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AsyncInfo As IAsyncInfo, m_AsyncOperation As IUnknown, m_sAsyncOperationName As String, AsyncOperation_PutCompleted As Long, AsyncOperation_GetResults As Long, _
        m_hEvent As LongPtr, m_cyStart As Currency

Friend Function Await(AsyncInfo As IAsyncInfo, Optional ByVal cyTimeout As Currency = INFINITE) As AsyncStatus
Dim vParam As Variant, hRes As Variant
    If Not IsBusy Then ' Guard against undesired reentrancy
        If AsyncOperation(AsyncInfo) Then
            hRes = S_FALSE: If m_hEvent = vbNullPtr Then m_hEvent = CreateEventW ' Create an auto-reset event object (resets automatically after being signaled)
            Set vParam = NewAsyncOperationCompletedHandler(m_hEvent) ' Create our light-weight Completed Handler delegate for this AsyncOperation
            DispCallFunc ObjPtr(m_AsyncOperation), AsyncOperation_PutCompleted, CC_STDCALL, vbLong, 1, VarType(vParam), VarPtr(vParam), hRes ' Equivalent to: Set AsyncOperation.Completed = vParam
            With m_AsyncInfo
                If hRes = S_OK Then
                    If cyTimeout <> INFINITE Then StartTiming m_cyStart
                    Do ' Sometimes the AsyncOperation will have already completed before starting the loop but we need to wait for the delegate handler to be executed and signal the event object!
                        Select Case MsgWaitForMultipleObjects(1, VarPtr(m_hEvent), dwMilliseconds:=cyTimeout)
                            Case WAIT_OBJECT_0: Exit Do ' The event has been signaled confirming the completion of the AsyncOperation
                            Case WAIT_OBJECT_0 + 1: DoEvents ' One or more messages are waiting in the queue and need to be processed
                            Case WAIT_TIMEOUT: .Cancel ' If a timeout has been provided then cancel the AsyncOperation (the completion handler is still executed thus signaling the event)
                        End Select
                        If cyTimeout <> INFINITE Then cyTimeout = cyTimeout - Elapsed(m_cyStart): StartTiming m_cyStart: If cyTimeout < 1 Then cyTimeout = 1
                    Loop
                End If
                Await = .Status
                #If bInIDE Then
                    If .Status = AsyncStatus_Error Then Debug.Print m_sAsyncOperationName, Hex$(.ErrorCode), GetErrorMessage(.ErrorCode) ': Debug.Assert False
                #End If
            End With
        End If
    End If
End Function

Friend Sub CancelAsyncOperation()
    If IsBusy Then m_AsyncInfo.Cancel
End Sub

Friend Property Get ErrorCode() As Long
    If Not m_AsyncInfo Is Nothing Then ErrorCode = m_AsyncInfo.ErrorCode Else ErrorCode = E_POINTER
End Property

Friend Property Get ErrorDescription() As String
    ErrorDescription = GetErrorMessage(ErrorCode)
End Property

Friend Function GetResults(Optional ByVal GetResultsPrimitive As LongPtr) As IUnknown
Dim vParam As Variant
    If Not m_AsyncOperation Is Nothing Then
        If GetResultsPrimitive = vbNullPtr Then vParam = VarPtr(GetResults) Else vParam = GetResultsPrimitive ' An AsyncOperation can return either an object or a primitive type
        DispCallFunc ObjPtr(m_AsyncOperation), AsyncOperation_GetResults, CC_STDCALL, vbLong, 1, VarType(vParam), VarPtr(vParam), vParam ' Call AsyncOperation.GetResults
        #If bInIDE Then
            If vParam <> S_OK Then Debug.Print m_sAsyncOperationName, Hex$(vParam), GetErrorMessage(vParam) ': Debug.Assert False
        #End If
    End If
End Function

Friend Property Get IsBusy() As Boolean
    IsBusy = Status = AsyncStatus_Started
End Property

Friend Property Get Name() As String
    Name = m_sAsyncOperationName
End Property

Friend Property Get Status() As AsyncStatus
    If Not m_AsyncInfo Is Nothing Then Status = m_AsyncInfo.Status Else Status = E_POINTER
End Property

Private Function AsyncOperation(AsyncInfo As IAsyncInfo) As Boolean ' All AsyncOperations implement IAsyncInfo
Dim vParam As Variant, hRes As Variant, InterfaceIIDs() As UUID, i As Long
    If Not m_AsyncInfo Is Nothing Then m_AsyncInfo.Close
    Set m_AsyncInfo = AsyncInfo
    If Not m_AsyncInfo Is Nothing Then
        m_sAsyncOperationName = IInspectableGetRuntimeClassName(m_AsyncInfo) ' Retrieve the verbose name of this AsyncOperation
        If InStr(m_sAsyncOperationName, "WithProgress") = 0 Then ' AsyncOperation and AsyncOperationWithProgress have different method offsets
            AsyncOperation_PutCompleted = 6 * PTR_SIZE: AsyncOperation_GetResults = 8 * PTR_SIZE
        Else
            AsyncOperation_PutCompleted = 8 * PTR_SIZE: AsyncOperation_GetResults = 10 * PTR_SIZE
        End If
        If IInspectableGetIIDs(m_AsyncInfo, InterfaceIIDs) Then ' Find the correct IID of this AsyncOperation and cast to it from IAsyncInfo
            For i = LBound(InterfaceIIDs) To UBound(InterfaceIIDs)
                If InterfaceIIDs(i).Data2 Then Set m_AsyncOperation = Nothing: AsyncOperation = QueryObj(ObjPtr(m_AsyncInfo), VarPtr(InterfaceIIDs(i)), VarPtr(m_AsyncOperation)) = S_OK: Exit Function
            Next i
        End If
    End If
End Function

Private Sub Class_Terminate()
    CloseHandle m_hEvent: If Not m_AsyncInfo Is Nothing Then m_AsyncInfo.Close
End Sub
