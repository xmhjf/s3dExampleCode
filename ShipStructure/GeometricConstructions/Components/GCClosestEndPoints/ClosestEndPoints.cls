VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClosestEndPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Curve1", "Select a curve", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Curve2", "Select a curve", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddOutput(1, "Point1")
    Call pGeometricConstructionDefinition.AddOutput(1, "Point2")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim dDistanceMini As Double: Let dDistanceMini = 100000
    Dim pPositionOfClosestEndPoint1 As IJDPosition
    Dim pPositionOfClosestEndPoint2 As IJDPosition
    Dim i As Integer
    For i = 0 To 1
        Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromWireBody(ModelBody_FromGeometry(pGeometricConstruction.Inputs("Curve1").Item(1)), i)
        
        Dim j As Integer
        For j = 0 To 1
            Dim pPosition2 As IJDPosition:: Set pPosition1 = Position_FromWireBody(ModelBody_FromGeometry(pGeometricConstruction.Inputs("Curve2").Item(1)), j)
            Dim dDistance As Double: Let dDistance = pPosition1.DistPt(pPosition2)
            If dDistance < dDistanceMini Then
                Set pPositionOfClosestEndPoint1 = pPosition1
                Set pPositionOfClosestEndPoint2 = pPosition2
                
                dDistanceMini = dDistance
            End If
        Next
    Next
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Point1", 1) = Point_FromPosition(pPositionOfClosestEndPoint1)
    oGCMacro.Output("Point2", 1) = Point_FromPosition(pPositionOfClosestEndPoint2)
End Sub
