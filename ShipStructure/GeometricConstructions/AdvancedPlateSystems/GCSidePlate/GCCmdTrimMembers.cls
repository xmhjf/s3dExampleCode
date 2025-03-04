VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCCmdTrimMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("PlateSystem", "Select a Side-PlateSystem", "IJPlateSystem", 1, 1)
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction)
    Dim oPlateSystem As Object
    Set oPlateSystem = pGeometricConstruction.Input("PlateSystem")
    
    Dim pNamedItem As IJNamedItem
    Set pNamedItem = oPlateSystem
    Debug.Print "Name= " + pNamedItem.Name
    
    Dim oSupport As Object
    ' Set oSupport = PlateSystem_GetSupport(oPlateSystem)
    
    Dim pGeometricConstruction1 As IJGeometricConstruction
    Set pGeometricConstruction1 = Entity_GetElementsOfRelatedEntities(oSupport, "IJGeometry", "ConstructionForOutput").Item(1)
    Debug.Print "Type= " + pGeometricConstruction1.Type
    
    Dim pGeometricConstructionMacro1 As IJGeometricConstructionMacro
    Set pGeometricConstructionMacro1 = pGeometricConstruction1

    ' initialize common node
    Dim pPositionOfNode As IJDPosition
'    Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = Position_GetCommonExtremityOfLines(pLinesOfMembers(1), pLinesOfMembers(2))
    If pPositionOfNode Is Nothing Then
        MsgBox "Common node not found"
    End If
        
'''    Call Members_TrimAtNode(pGeometricConstruction1.Inputs("MemberParts"), _
'''                            pGeometricConstructionMacro1.Outputs("Boundary"), _
'''                            oSupport, _
'''                            pPositionOfNode)
End Sub

