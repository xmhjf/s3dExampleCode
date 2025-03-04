VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectPointType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Implements IJGCTypeInfo
Implements IJGeometricConstructionDefinitionService

'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "Select point type"
'End Property
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "SelectPointType"
'End Property
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Connectable1", "Select a root platesystem", "IJPlateSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootPlateSystem]", 0, 1, "IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod")
    Call pGeometricConstructionDefinition.AddInput("Connectable2", "Select a root profile system", "IJStiffenerSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootStiffenerSystem]", 0, 1, "IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod")
    
    Call pGeometricConstructionDefinition.AddParameter("Choice", "Choice", GCLong, 0, 0, 0, 0, 0, True)
    
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCGTypePoint3d, "Point")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCLocalCoordinateSystem, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As RESPOMLib.IJDPOM)
    ' prepare result
    Dim iChoice As Integer: Let iChoice = 0
    
    ' retrieve inputs
    Dim oConnectable1 As Object: Set oConnectable1 = GeometricConstruction_GetInput(pGC, "Connectable1")
    Dim oConnectable2 As Object: Set oConnectable2 = GeometricConstruction_GetInput(pGC, "Connectable2")

    If Not oConnectable1 Is Nothing _
    And Not oConnectable2 Is Nothing Then
        Let iChoice = 1
        If TypeOf oConnectable1 Is IJPlateSystem _
        And TypeOf oConnectable2 Is IJStiffenerSystem Then
            Dim pDesignChild As IJDesignChild
            Set pDesignChild = oConnectable2
            If pDesignChild.GetParent Is oConnectable1 Then Let iChoice = 2
        End If
    End If

    ' set result
    If pGC.Parameter("Choice") <> iChoice Then pGC.Parameter("Choice") = iChoice
End Sub

