VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStorageItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IIterable_IStorageItem
Implements IIterator_IStorageItem

Private BackingStorage As New Collection, m_CurrentIndex As Long

Friend Function This(vFiles As Variant) As cStorageItems
Dim vFile As Variant
    If Not IsArray(vFiles) Then vFiles = Array(vFiles)
    With New cAwait
        For Each vFile In vFiles
            If .Await(StorageFileStatics.GetFileFromPathAsync(StrRef(vFile))) = AsyncStatus_Completed Then BackingStorage.Add .GetResults
        Next vFile
    End With
    Set This = Me
End Function

Private Function IIterable_IStorageItem_First() As IIterator_IStorageItem
    m_CurrentIndex = 0: Set IIterable_IStorageItem_First = Me
    Debug.Print "IIterable_IStorageItem_First"
End Function

Private Property Get IIterator_IStorageItem_Current() As IStorageItem
    Set IIterator_IStorageItem_Current = BackingStorage(m_CurrentIndex + 1)
    Debug.Print "IIterator_IStorageItem_Current"
End Property

Private Property Get IIterator_IStorageItem_HasCurrent() As BooleanByte
    IIterator_IStorageItem_HasCurrent = Abs(m_CurrentIndex <= BackingStorage.Count - 1)
    Debug.Print "IIterator_IStorageItem_HasCurrent"
End Property

Private Function IIterator_IStorageItem_MoveNext() As BooleanByte
    m_CurrentIndex = m_CurrentIndex + 1: IIterator_IStorageItem_MoveNext = IIterator_IStorageItem_HasCurrent
    Debug.Print "IIterator_IStorageItem_MoveNext"
End Function

Private Function IIterator_IStorageItem_GetMany(ByVal Length As Long, ByVal Items As LongPtr) As Long
    ' Not Implemented
End Function
