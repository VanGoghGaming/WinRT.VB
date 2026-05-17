VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cBasicProperties
Implements IAsyncOperationCompletedHandlerIStorageFolder
Implements IAsyncOperationCompletedHandlerUInt32
Implements IAsyncOperationCompletedHandlerIVectorView_IStorageFile
Implements ITypedEventHandlerStorageQueryResultBaseIInspectable

Event FileCreated(ByVal sPath As String)
Event FileDeleted(ByVal sPath As String)
Event FileChanged(ByVal sPath As String, ByVal FileChange As FileChanges, ByVal FileSize As UInt64, ByVal OldSize As UInt64, ByVal DateModified As Date, ByVal OldDate As Date, ByVal FileAttributes As FileAttributes, ByVal OldAttributes As FileAttributes)

Private m_FileQuery As IStorageFileQueryResult, m_StorageFolderQueryOperations As IStorageFolderQueryOperations, m_Files As Collection, m_FileCount As Long

Friend Property Get FileCount() As Long
    FileCount = m_FileCount
End Property

Friend Function This(sWatchedFolder As String) As cFileQuery
    Set StorageFolderStatics.GetFolderFromPathAsync(StrRef(sWatchedFolder)).Completed = Me: Set This = Me
End Function

Private Sub cBasicProperties_FileChanged(sPath As String, ByVal FileChange As FileChanges, ByVal FileSize As UInt64, ByVal OldSize As UInt64, ByVal DateModified As Date, ByVal OldDate As Date, ByVal FileAttributes As FileAttributes, ByVal OldAttributes As FileAttributes)
    RaiseEvent FileChanged(sPath, FileChange, FileSize, OldSize, DateModified, OldDate, FileAttributes, OldAttributes)
End Sub

Private Sub EnumerateFiles(Optional ByVal bCheckDeletedFiles As Boolean)
Dim i As Long, BasicProperties As cBasicProperties
    For i = m_Files.Count To 1 Step -1
        Set BasicProperties = m_Files(i)
        With BasicProperties
            If bCheckDeletedFiles Then
                If Not .FileExists Then m_Files.Remove i: RaiseEvent FileDeleted(.Path)
            Else
                .FileExists = False ' Reset state
            End If
        End With
    Next i
End Sub

Private Function HasKey(sKey As String, BasicProperties As cBasicProperties) As Boolean
On Error GoTo EH
    Set BasicProperties = m_Files(sKey): HasKey = True: Exit Function
EH: Err.Clear
End Function

Private Sub IAsyncOperationCompletedHandlerIStorageFolder_Invoke(ByVal AsyncOperation As IAsyncOperationIStorageFolder, ByVal AsyncStatus As AsyncStatus)
Dim AsyncInfo As IAsyncInfo, StorageQueryResultBase As IStorageQueryResultBase
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            Set m_StorageFolderQueryOperations = AsyncOperation.GetResults
            Set m_FileQuery = m_StorageFolderQueryOperations.CreateFileQueryOverloadDefault
            Set StorageQueryResultBase = m_FileQuery: StorageQueryResultBase.AddContentsChanged Me
            Set m_FileQuery.GetFilesAsyncDefaultStartAndCount.Completed = Me
        Case AsyncStatus_Error
            Set AsyncInfo = AsyncOperation
            With AsyncInfo
                Debug.Print Hex$(.ErrorCode), GetErrorMessage(.ErrorCode)
            End With
    End Select
End Sub

Private Sub IAsyncOperationCompletedHandlerIVectorView_IStorageFile_Invoke(ByVal AsyncOperation As IAsyncOperationIVectorView_IStorageFile, ByVal AsyncStatus As AsyncStatus)
Dim AsyncInfo As IAsyncInfo, StorageItem As IStorageItem, i As Long, sPath As String, bNewCollection As Boolean, BasicProperties As cBasicProperties
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            If m_Files Is Nothing Then Set m_Files = New Collection: bNewCollection = True
            EnumerateFiles
            With AsyncOperation.GetResults
                For i = 0 To .Size - 1
                    Set StorageItem = .GetAt(i): sPath = GetStr(StorageItem.Path)
                    If bNewCollection Then
                        With New cBasicProperties: m_Files.Add .This(Me, sPath, StorageItem), sPath: End With
                    Else
                        If HasKey(sPath, BasicProperties) Then
                            BasicProperties.GetBasicProperties StorageItem
                        Else
                            With New cBasicProperties: m_Files.Add .This(Me, sPath, StorageItem), sPath: End With
                            RaiseEvent FileCreated(sPath)
                        End If
                    End If
                Next i
            End With
            EnumerateFiles True
        Case AsyncStatus_Error
            Set AsyncInfo = AsyncOperation
            With AsyncInfo
                Debug.Print Hex$(.ErrorCode), GetErrorMessage(.ErrorCode)
            End With
    End Select
End Sub

Private Sub IAsyncOperationCompletedHandlerUInt32_Invoke(ByVal AsyncOperation As IAsyncOperationUInt32, ByVal AsyncStatus As AsyncStatus)
    Select Case AsyncStatus
        Case AsyncStatus_Completed: m_FileCount = AsyncOperation.GetResults
    End Select
End Sub

Private Sub ITypedEventHandlerStorageQueryResultBaseIInspectable_Invoke(ByVal Sender As IStorageQueryResultBase, ByVal Args As IInspectable)
Dim FileQuery As IStorageFileQueryResult
    Set Sender.GetItemCountAsync.Completed = Me
    Set FileQuery = Sender: Set FileQuery.GetFilesAsyncDefaultStartAndCount.Completed = Me
End Sub
