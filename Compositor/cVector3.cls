VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVector3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_X As Single, m_Y As Single, m_Z As Single

Friend Function This(Optional ByVal X As Single, Optional ByVal Y As Single, Optional ByVal Z As Single) As cVector3
    m_X = X: m_Y = Y: m_Z = Z: Set This = Me
End Function

Friend Property Get V3() As Vector3
    With V3: .X = m_X: .Y = m_Y: .Z = m_Z: End With
End Property
