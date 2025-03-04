VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'************************************************************************
'  Method : GetPI - returns PI value
'************************************************************************
Public Function GetPI() As Double
    GetPI = PI
End Function

Public Sub ForceUpdateSmartItem(oObject As Object)

On Error GoTo ErrorHandler
Const sMETHOD = "ForceUpdateSmartItem"
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute WebCut before FlangeCuts
    ' interface IJStructGeometry : {6034AD40-FA0B-11d1-B2FD-080036024603}
    Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools

    Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
    oStructAssocTools.UpdateObject oObject, _
                                   "{6034AD40-FA0B-11d1-B2FD-080036024603}"
    Set oStructAssocTools = Nothing

    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, "RuleKeys", sMETHOD).Number

End Sub

