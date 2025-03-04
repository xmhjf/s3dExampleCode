VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSAsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements IJGCTypeInfo
Implements IJGeometricConstructionDefinitionService
Private Const sCOORDINATE_SYSTEM1 = "CoordinateSystem1"
Private Const sCOORDINATE_SYSTEM = "CoordinateSystem"
'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "CS as CS"
'End Property
'
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "TestCSAsCS"
'End Property

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput(sCOORDINATE_SYSTEM1, "Select a coordinate system", "IJDCoordinateSystem", 0, 1)
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCLocalCoordinateSystem, sCOORDINATE_SYSTEM)
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oCS As SP3DGeometricConstruction.GeometricConstruction
    If pGC.Inputs(sCOORDINATE_SYSTEM1).Count = 1 Then
        Set oCS = oGCFactory.CreateEntity("CSByCS", Nothing)
        oCS.Input(sCOORDINATE_SYSTEM) = pGC.Input(sCOORDINATE_SYSTEM1)
    Else
        Set oCS = oGCFactory.CreateEntity("CSAsRoot", Nothing)
    End If
    oCS.Evaluate

    Dim pGCMacro As IJGeometricConstructionMacro
    Set pGCMacro = pGC
    pGCMacro.Output(sCOORDINATE_SYSTEM, 1) = oCS
End Sub

