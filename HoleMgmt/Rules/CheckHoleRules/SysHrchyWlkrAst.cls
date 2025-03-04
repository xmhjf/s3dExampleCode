VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysHrchyWlkrAst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements GSCADHierarchyWalker.IJHierarchyWalkerAsst

Private m_oDPom As IJDPOM 'to check interface support

Private Sub Class_Terminate()
On Error GoTo OnErr

Set m_oDPom = Nothing

Exit Sub
OnErr:
    MsgBox "Error: " & Err.Number & "," & vbCrLf & "Class_Terminate: " _
        & Err.Description, vbOKOnly + vbCritical, Err.Source
End Sub

Private Property Let IJHierarchyWalkerAsst_DPom(ByVal RHS As GSCADHierarchyWalker.IJDPOM)
    Set m_oDPom = RHS
End Property

Private Function IJHierarchyWalkerAsst_Continue(ByVal pMonikerUnk As Variant) As Boolean
On Error GoTo OnErr
    IJHierarchyWalkerAsst_Continue = False 'default

    'check for support of IJSystem
    If m_oDPom.SupportsInterface(pMonikerUnk, "IJSystem") Then
        IJHierarchyWalkerAsst_Continue = True
    End If

Exit Function
OnErr:
    MsgBox "Error: " & Err.Number & "," & vbCrLf & "IJHierarchyWalkerAsst_Continue: " _
        & Err.Description, vbOKOnly + vbCritical, Err.Source

End Function

Private Function IJHierarchyWalkerAsst_Keep(ByVal pMonikerUnk As Variant) As Boolean
On Error GoTo OnErr

    IJHierarchyWalkerAsst_Keep = False 'default

    'check for support of IJSystem
    If m_oDPom.SupportsInterface(pMonikerUnk, "IJSystemChild") <> 0 Then
        IJHierarchyWalkerAsst_Keep = True
    End If

Exit Function
OnErr:
    MsgBox "Error: " & Err.Number & "," & vbCrLf & "IJHierarchyWalkerAsst_Keep: " _
        & Err.Description, vbOKOnly + vbCritical, Err.Source
End Function

 