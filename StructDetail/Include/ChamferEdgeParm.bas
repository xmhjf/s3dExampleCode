Attribute VB_Name = "ChamferEdgeParm"
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\ChamferEdgeParm.bas"
'

'***********************************************************************
' METHOD:  PlateEdgeByStiffenerEdgeChamferData
'
' DESCRIPTION:
'
'***********************************************************************
Public Sub PlateEdgeByStiffenerEdgeChamferData(oPlatePort1 As Object, _
                                               oStiffenerPort2 As Object, _
                                               dChamferBase As Double, _
                                               dChamferOffset As Double)
                                                              
Const METHOD = "::PlateEdgeByStiffenerEdgeChamferData"
On Error GoTo ErrorHandler
    
    Dim oSDOex_Chamfer As StructDetailObjectsex.Chamfer
    
    dChamferBase = 0#
    dChamferOffset = 0#
    
    Set oSDOex_Chamfer = New StructDetailObjectsex.Chamfer
    oSDOex_Chamfer.GetStiffToPlateEdgeChamferData oPlatePort1, oStiffenerPort2, dChamferBase, dChamferOffset
    
    Exit Sub
    
ErrorHandler:
    LogError Err, MODULE, METHOD
    Err.Clear
End Sub







