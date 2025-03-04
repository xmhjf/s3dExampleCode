VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectFreeEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Implements IJGCTypeInfo
Implements IJGeometricConstructionDefinitionService
'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "Select free edge"
'End Property
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "SelectFreeEdge"
'End Property
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddParameter("Choice", "Choice", GCLong, 0, 0, 0, 0, 0, True)
    
    Call pGeometricConstructionDefinition.AddParameter("FreeEdge", "FreeEdge", GCCodeList, 0, 0, 0, 0, 1, True)
    Call pGeometricConstructionDefinition.AddParameterValue("FreeEdge", "Line by anchor points", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("FreeEdge", "Curve at offset", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("FreeEdge", "Curve at offset with arcs to anchor points", 3)
    
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As RESPOMLib.IJDPOM)
    ' prepare result
    Dim iChoice As Integer: Let iChoice = 0
    
    Let iChoice = pGC.Parameter("FreeEdge")
    
    ' set result
    If CLng(pGC.Parameter("Choice")) <> iChoice Then pGC.Parameter("Choice") = iChoice
End Sub


