VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBindableRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cRecord
Implements IBindableIterable
Implements IBindableIterator
Implements IBindableVector
Implements IWeakRefSrc
Implements Unrestricted.IInspectable

Private m_CurrentIndex As Long, m_RecordSet As ADODB.RecordSet, m_WeakReference As IWeakRef

Friend Function This(RecordSet As ADODB.RecordSet) As cBindableRecordSet
    Set m_RecordSet = RecordSet: Set This = Me ' This constructor saves a local instance of the ADODB RecorSet
End Function

Private Function GetRecord(ByVal Index As Long) As IMap_IInspectable
Dim Field As ADODB.Field, ObservableMap As IObservableMap_IInspectable
    Set GetRecord = NewObject("PropertySet") ' Here we instantiate the PropertySet class (which is an observable collection of objects with string keys)
    ' The PropertySet class implements the IMap_IInspectable and IObservableMap_IInspectable interfaces (among others)
    m_RecordSet.AbsolutePosition = Index ' Select the appropriate row from our ADODB RecordSet
    For Each Field In m_RecordSet.Fields ' Iterate through the fields from the selected row and insert them in our collection (the field name serves as the key)
        GetRecord.Insert StrRef(Field.Name), Box(Field.Value) ' The collection wants objects so we need to box the field values (numbers, strings, etc) into objects
    Next Field
    ' The cRecord class serves as the handler for receiving Changed notifications from our observable collection above:
    Set ObservableMap = GetRecord
    With New cRecord
        ObservableMap.AddMapChanged .This(Me, Index) ' Subscribe to the MapChanged event and pass the necessary parameters so we know which row has been changed
    End With
End Function

Private Sub cRecord_MapChanged(ByVal Index As Long, ByVal Sender As IMap_IInspectable, ByVal Args As IMapChangedEventArgs) ' We receive this notification when editing a TextBox from the ListView
Dim sKey As String, Value As Variant
    sKey = GetStr(Args.Key): Value = Unbox(Sender.Lookup(StrRef(sKey))) ' Retrieve the name and value of the current field being edited (the value is the object we boxed earlier so we need to unbox it now)
    Debug.Print Choose(Args.ChangeType + 1, "CollectionReset", "ItemInserted", "ItemRemoved", "ItemChanged"), sKey, Value
    If Value <> vbNullString Then ' This database does not allow empty strings
        With m_RecordSet
            .AbsolutePosition = Index: .Fields(sKey) = Value: .Update ' Save any changes immediately (usually some sort of validation should be performed first)
        End With
    End If
    frmBindableRecordset.Caption = "Changing row " & Index & " - " & sKey & " - " & Value
End Sub

Private Function IBindableIterable_First() As IBindableIterator ' Boilerplate code (doesn't get called but we are required to implement IBindableIterable and IBindableIterator)
    m_CurrentIndex = 1: Set IBindableIterable_First = Me
    #If bInIDE Then
        Debug.Print "IBindableIterable_First"
    #End If
End Function

Private Property Get IBindableIterator_Current() As IInspectable
    Set IBindableIterator_Current = GetRecord(m_CurrentIndex)
    #If bInIDE Then
        Debug.Print "IBindableIterator_Current"
    #End If
End Property

Private Property Get IBindableIterator_HasCurrent() As BooleanByte
    Select Case m_CurrentIndex
        Case 1 To m_RecordSet.RecordCount: IBindableIterator_HasCurrent = APITRUE
    End Select
End Property

Private Function IBindableIterator_MoveNext() As BooleanByte
    m_CurrentIndex = m_CurrentIndex + 1: IBindableIterator_MoveNext = IBindableIterator_HasCurrent
    #If bInIDE Then
        Debug.Print "IBindableIterator_MoveNext"
    #End If
End Function

Private Sub IBindableVector_Append(ByVal Value As IInspectable)
    #If bInIDE Then
        Debug.Print "IBindableVector_Append"
    #End If
End Sub

Private Sub IBindableVector_Clear()
    #If bInIDE Then
        Debug.Print "IBindableVector_Clear"
    #End If
End Sub

Private Function IBindableVector_GetAt(ByVal Index As Long) As IInspectable ' The ListView control will call this method a lot when populating its rows with values from the RecordSet
    Select Case Index + 1 ' The AbsolutePosition method of the RecordSet is one-based so we need to adjust the index
        Case 1 To m_RecordSet.RecordCount: Set IBindableVector_GetAt = GetRecord(Index + 1) ' Retrieve the requested row from the ADODB RecordSet and copy it in an observable collection
    End Select
    #If bInIDE Then
'        Debug.Print "IBindableVector_GetAt", Index
    #End If
End Function

Private Function IBindableVector_GetView() As IBindableVectorView
    #If bInIDE Then
        Debug.Print "IBindableVector_GetView"
    #End If
End Function

Private Function IBindableVector_IndexOf(ByVal Value As IInspectable, Index As Long) As BooleanByte
    #If bInIDE Then
        Debug.Print "IBindableVector_IndexOf"
    #End If
End Function

Private Sub IBindableVector_InsertAt(ByVal Index As Long, ByVal Value As IInspectable)
    #If bInIDE Then
        Debug.Print "IBindableVector_InsertAt"
    #End If
End Sub

Private Sub IBindableVector_RemoveAt(ByVal Index As Long)
    #If bInIDE Then
        Debug.Print "IBindableVector_RemoveAt"
    #End If
End Sub

Private Sub IBindableVector_RemoveAtEnd()
    #If bInIDE Then
        Debug.Print "IBindableVector_RemoveAtEnd"
    #End If
End Sub

Private Sub IBindableVector_SetAt(ByVal Index As Long, ByVal Value As IInspectable)
    #If bInIDE Then
        Debug.Print "IBindableVector_SetAt"
    #End If
End Sub

Private Property Get IBindableVector_Size() As Long ' This method is also called by the ListView control to check the total number of rows
    IBindableVector_Size = m_RecordSet.RecordCount
    #If bInIDE Then
'        Debug.Print "IBindableVector_Size"
    #End If
End Property

Private Sub IInspectable_GetIIDs(iidCount As Long, iidPtr As LongPtr)
    #If bInIDE Then
        Debug.Print "IInspectable_GetIIDs"
    #End If
End Sub

Private Function IInspectable_GetRuntimeClassName() As HSTRING
    #If bInIDE Then
        Debug.Print "IInspectable_GetRuntimeClassName"
    #End If
End Function

Private Function IInspectable_GetTrustLevel() As TrustLevel
    #If bInIDE Then
        Debug.Print "IInspectable_GetTrustLevel"
    #End If
End Function

Private Function IWeakRefSrc_GetWeakRef() As IWeakRef
    If m_WeakReference Is Nothing Then Set m_WeakReference = NewWeakRef(ObjPtr(Me))
    Set IWeakRefSrc_GetWeakRef = m_WeakReference
End Function

Private Sub Class_Terminate()
    If Not m_WeakReference Is Nothing Then m_WeakReference.StillAlive = False
    If Not m_RecordSet Is Nothing Then If (m_RecordSet.State And adStateOpen) = adStateOpen Then m_RecordSet.Close ' Make sure the database file is closed properly
    #If bInIDE Then
        Debug.Print "Class_Terminate cBindableRecordSet"
    #End If
End Sub
