VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWalkerAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJHierarchyWalkerAsst

Private m_POM As IJDPOM

Private Sub Class_Initialize()
    Set m_POM = Nothing
End Sub

Private Function IJHierarchyWalkerAsst_Continue(ByVal pMoniker As Variant) As Boolean
    IJHierarchyWalkerAsst_Continue = True
End Function

Private Property Let IJHierarchyWalkerAsst_DPom(ByVal RHS As GSCADHierarchyWalker.IJDPOM)
    Set m_POM = RHS
End Property

Private Function IJHierarchyWalkerAsst_Keep(ByVal pMoniker As Variant) As Boolean

    IJHierarchyWalkerAsst_Keep = False
    
    If m_POM.SupportsInterface(pMoniker, "IJMfgProfilePart") Then
        IJHierarchyWalkerAsst_Keep = True
    End If

End Function
