VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMapChangedEventHandlerIMapIInspectable ' This handler fires when the observable collection has changed in any way (Item Inserted, Removed, Changed)

Private m_Index As Long, m_Callback As IWeakRef

Friend Function This(Callback As IWeakRefSrc, ByVal Index As Long) As cRecord
    Set m_Callback = Callback.GetWeakRef: m_Index = Index: Set This = Me ' Save the index of this row as well as the Callback object used to forward notifications.
End Function

Public Sub MapChanged(ByVal Index As Long, ByVal Sender As IMap_IInspectable, ByVal Args As IMapChangedEventArgs)
' This method is declared as Public because it will be called from parent objects that implement this one
End Sub

Private Sub IMapChangedEventHandlerIMapIInspectable_Invoke(ByVal Sender As IObservableMap_IInspectable, ByVal Args As IMapChangedEventArgs)
Dim Record As cRecord
    If Not m_Callback Is Nothing Then
        If m_Callback.StillAlive Then Set Record = m_Callback.Resolve: Record.MapChanged m_Index, Sender, Args ' Notify the parent object which record has been changed
    End If
End Sub
