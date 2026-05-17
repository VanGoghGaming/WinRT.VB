VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFilePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FilePickerResult
    FilePickerResult_OK
    FilePickerResult_Canceled
    FilePickerResult_Timeout
    FilePickerResult_Error
End Enum

Private m_FileOpenPicker As IFileOpenPicker, m_FileName As String, m_FilePath As String

Friend Property Get FileName() As String
    FileName = m_FileName
End Property

Friend Property Get FullPath() As String
    FullPath = m_FilePath
End Property

Friend Sub SetFileTypeFilters(ParamArray vFilters() As Variant)
Dim vFilter As Variant
    With m_FileOpenPicker.FileTypeFilter
        .Clear
        If IsMissing(vFilters) Then
            .Append StrRef("*")
        Else
            For Each vFilter In vFilters
                If VarType(vFilter) = vbString Then .Append StrRef(vFilter)
            Next vFilter
        End If
    End With
End Sub

Friend Function SetOwnerWindow(hWndOwner As LongPtr) As Boolean
    SetOwnerWindow = InitializeWithWindow(m_FileOpenPicker, hWndOwner)
End Function

Friend Function ShowFileOpenPicker(StorageItem As IStorageItem, Optional dwMilliseconds As Long = INFINITE) As FilePickerResult
    With New cAwait
        Select Case .Await(m_FileOpenPicker.PickSingleFileAsync, dwMilliseconds)
            Case AsyncStatus_Completed
                Set StorageItem = .GetResults
                If Not StorageItem Is Nothing Then
                    ShowFileOpenPicker = FilePickerResult_OK: m_FileName = GetStr(StorageItem.Name): m_FilePath = GetStr(StorageItem.Path)
                Else
                    ShowFileOpenPicker = FilePickerResult_Canceled
                End If
            Case AsyncStatus_Canceled: ShowFileOpenPicker = FilePickerResult_Timeout
            Case AsyncStatus_Error: ShowFileOpenPicker = FilePickerResult_Error
        End Select
    End With
End Function

Private Sub Class_Initialize()
    Set m_FileOpenPicker = NewObject("FileOpenPicker")
End Sub

