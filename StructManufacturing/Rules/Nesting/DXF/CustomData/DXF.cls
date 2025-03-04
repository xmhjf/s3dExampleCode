VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DXF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJMfgOutputData

Private Function IJMfgOutputData_Generate(ByVal pMfgObj As Object) As Object
    Dim oMfgPlate As IJMfgPlatePart
    Set oMfgPlate = pMfgObj

    If Not oMfgPlate Is Nothing Then
        Set IJMfgOutputData_Generate = oMfgPlate.GetXMLData
    End If

End Function

Private Sub IJMfgOutputData_Modify(ByVal pMfgObj As Object, ppDoc As Object)

'    TR-CP-227633  Few delivered MFG rules cannot be recompiled on an End user machine
'    Removing the functionality since IJMfgXMLData is changed recently.


'    Dim oAdaptor As IJMfgXMLData
'    Set oAdaptor = SP3DCreateObject("MfgXMLAdaptor.GNest")
'    oAdaptor.Update pMfgObj, ppDoc
End Sub
