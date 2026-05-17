VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBasicProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAsyncOperationCompletedHandlerIBasicProperties

Public Enum FileChanges
    FileChanges_None
    FileChanges_SizeModified
    FileChanges_DateModified
    FileChanges_AttributesModified = 4
End Enum

Private m_Callback As cBasicProperties, m_sPath As String, m_StorageItem As IStorageItem, m_bFileExists As Boolean, m_DateModified As WinRT.DateTime, m_FileSize As UInt64, m_FileAttributes As FileAttributes, _
        OldSize As UInt64, OldDate As WinRT.DateTime, OldAttributes As FileAttributes

Friend Property Get FileExists() As Boolean
    FileExists = m_bFileExists
End Property

Friend Property Let FileExists(ByVal bFileExists As Boolean)
    m_bFileExists = bFileExists
End Property

Friend Sub GetBasicProperties(StorageItem As IStorageItem)
    m_bFileExists = True: Set m_StorageItem = StorageItem: Set m_StorageItem.GetBasicPropertiesAsync.Completed = Me
End Sub

Friend Property Get Path() As String
    Path = m_sPath
End Property

Friend Function This(Callback As cBasicProperties, sPath As String, StorageItem As IStorageItem) As cBasicProperties
    Set m_Callback = Callback: m_sPath = sPath: m_FileSize = -1: m_FileAttributes = -1: GetBasicProperties StorageItem: Set This = Me
End Function

Public Sub FileChanged(sPath As String, ByVal FileChange As FileChanges, ByVal FileSize As UInt64, ByVal OldSize As UInt64, ByVal DateModified As Date, ByVal OldDate As Date, ByVal FileAttributes As FileAttributes, ByVal OldAttributes As FileAttributes)

End Sub

Private Sub IAsyncOperationCompletedHandlerIBasicProperties_Invoke(ByVal AsyncOperation As IAsyncOperationIBasicProperties, ByVal AsyncStatus As AsyncStatus)
Dim AsyncInfo As IAsyncInfo, FileChange As FileChanges
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            With AsyncOperation.GetResults
                If m_FileSize = -1 Then
                    m_FileSize = .Size: OldSize = m_FileSize
                ElseIf m_FileSize <> .Size Then
                    FileChange = FileChange Or FileChanges_SizeModified: OldSize = m_FileSize: m_FileSize = .Size
                End If
                If m_DateModified = 0 Then
                    m_DateModified = .DateModified: OldDate = m_DateModified
                ElseIf m_DateModified <> .DateModified Then
                    FileChange = FileChange Or FileChanges_DateModified: OldDate = m_DateModified: m_DateModified = .DateModified
                End If
                If m_FileAttributes = -1 Then
                    m_FileAttributes = m_StorageItem.Attributes: OldAttributes = m_FileAttributes
                ElseIf m_FileAttributes <> m_StorageItem.Attributes Then
                    FileChange = FileChange Or FileChanges_AttributesModified: OldAttributes = m_FileAttributes: m_FileAttributes = m_StorageItem.Attributes
                End If
                If FileChange Then m_Callback.FileChanged m_sPath, FileChange, m_FileSize * 10000, OldSize * 10000, ToDate(m_DateModified, True), ToDate(OldDate, True), m_FileAttributes, OldAttributes
            End With
        Case AsyncStatus_Error
            Set AsyncInfo = AsyncOperation
            With AsyncInfo
                Debug.Print Hex$(.ErrorCode), GetErrorMessage(.ErrorCode)
                Set m_StorageItem.GetBasicPropertiesAsync.Completed = Me
            End With
    End Select
End Sub
