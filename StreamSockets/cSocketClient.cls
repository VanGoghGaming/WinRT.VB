VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSocketClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAsyncOperationCompletedHandlerUInt32

Private m_SocketClient As IStreamSocket, m_Reader As IDataReader, m_Writer As IDataWriter, m_HostName As IHostName, m_Await As cAwait

Event DataReceived(sData As String)

Friend Function Connect(sHostName As String, ByVal sPort As String, Optional ByVal cyTimeout As Currency = 10000) As Boolean ' 10 seconds timeout
    With m_Await
        If Not .IsBusy Then
            Set m_HostName = HostNameFactory.CreateHostName(StrRef(sHostName))
            Select Case .Await(m_SocketClient.ConnectAsync(m_HostName, StrRef(sPort)), cyTimeout)
                Case AsyncStatus_Completed
                    Set m_Reader = DataReaderFactory.CreateDataReader(m_SocketClient.InputStream) ' Wrap the socket's InputStream into a friendlier DataReader object
                    m_Reader.InputStreamOptions = InputStreamOptions_Partial ' We need to be able to read chunks from the stream instead of everything at once
                    Set m_Reader.LoadAsync(1024).Completed = Me ' We'll receive a callback whenever there's data available to read from the stream
                    Set m_Writer = DataWriterFactory.CreateDataWriter(m_SocketClient.OutputStream) ' Wrap the socket's OutputStream into a friendlier DataWriter object
                    Connect = True
                Case AsyncStatus_Canceled
                    cPop.MessageBox "Connection attempt timed out!", App.Title, vbOK
                Case AsyncStatus_Error
                    cPop.MessageBox "Error: " & Hex$(.ErrorCode) & vbNewLine & .ErrorDescription, App.Title, vbOK
            End Select
        Else
            If cPop.MessageBox("Previous connection attempt still waiting, do you want to cancel it?", App.Title, vbYes, vbNo) = vbYes Then .CancelAsyncOperation
        End If
    End With
End Function

Friend Function UpgradeToSsl() As Boolean
    UpgradeToSsl = m_Await.Await(m_SocketClient.UpgradeToSslAsync(SocketProtectionLevel_Tls12, m_HostName)) = AsyncStatus_Completed
End Function

Friend Function SendMessage(sMessage As String) As Boolean
    With m_Await
        If Not .IsBusy And Not m_Writer Is Nothing Then
            m_Writer.WriteString StrRef(sMessage): SendMessage = .Await(m_Writer.StoreAsync) = AsyncStatus_Completed ' Calling StoreAsync is required to actually send the data over the network
        End If
    End With
End Function

Private Sub Class_Initialize()
    Set m_Await = New cAwait ' This simulates the functionality of the await syntax in .NET
    Set m_SocketClient = NewObject("StreamSocket") ' Create a new instance of the activatable StreamSocket Runtime Class
End Sub

Private Sub Class_Terminate()
    Dispose m_Reader: Dispose m_Writer: Dispose m_SocketClient ' These classes implement the IClosable interface so they need to be disposed of in order to free allocated resources
End Sub

Private Sub IAsyncOperationCompletedHandlerUInt32_Invoke(ByVal AsyncOperation As IAsyncOperationUInt32, ByVal AsyncStatus As AsyncStatus) ' This is the callback from the DataReader object reading data from the socket's InputStream
Dim lDataSize As Long, AsyncInfo As IAsyncInfo
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            lDataSize = AsyncOperation.GetResults ' When completed, this AsyncOperation returns how many bytes are available to read from the stream
            If lDataSize > 0 Then
                RaiseEvent DataReceived(GetStr(m_Reader.ReadString(lDataSize))) ' Read the current contents of the InputStream
                Set m_Reader.LoadAsync(1024).Completed = Me ' Set up the callback for the next time there's data available to read from the InputStream
            Else
                RaiseEvent DataReceived("Connection closed!") ' No more bytes to read means the server has closed the connection
            End If
        Case AsyncStatus_Error ' Just in case there are any errors we can see their description here
            Set AsyncInfo = AsyncOperation ' All AsyncOperations implement the IAsyncInfo interface
            With AsyncInfo
                RaiseEvent DataReceived(Hex$(.ErrorCode) & " : " & GetErrorMessage(.ErrorCode))
            End With
    End Select
End Sub
