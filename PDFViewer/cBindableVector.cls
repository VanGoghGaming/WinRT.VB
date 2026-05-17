VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBindableVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBindableIterable
Implements IBindableIterator
Implements IBindableVector
Implements Unrestricted.IInspectable

Private m_CurrentIndex As Long, m_PageCount As Long

Friend Function This(ByVal lPageCount As Long) As cBindableVector
    m_PageCount = lPageCount: Set This = Me
End Function

Private Function IBindableIterable_First() As IBindableIterator ' Boilerplate code (doesn't get called by the bound ListView control but we are required to implement IBindableIterable and IBindableIterator)
    m_CurrentIndex = 0: Set IBindableIterable_First = Me
    #If bInIDE Then
        Debug.Print "IBindableIterable_First"
    #End If
End Function

Private Property Get IBindableIterator_Current() As IInspectable
    Set IBindableIterator_Current = Box(m_CurrentIndex) ' Just return the current iteration index as a boxed integer
    #If bInIDE Then
        Debug.Print "IBindableIterator_Current"
    #End If
End Property

Private Property Get IBindableIterator_HasCurrent() As BooleanByte
    Select Case m_CurrentIndex
        Case 0 To m_PageCount - 1: IBindableIterator_HasCurrent = APITRUE
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

Private Function IBindableVector_GetAt(ByVal Index As Long) As IInspectable
    Set IBindableVector_GetAt = Box(Index) ' Just return the requested index as a boxed integer
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

Private Property Get IBindableVector_Size() As Long
    IBindableVector_Size = m_PageCount
    #If bInIDE Then
'        Debug.Print "IBindableVector_Size"
    #End If
End Property

Private Sub Class_Terminate()
    #If bInIDE Then
'        Debug.Print "Class_Terminate cBindableVector"
    #End If
End Sub

' Boilerplate IInspectable implementation (doesn't get called but our object needs to respond with S_OK to IInspectable queries)

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
