VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCCmdConnectMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("PlateSystem", "Select an Advanced PlateSystem", "IJPlateSystem", 1, 1)
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction)
    Dim pPlateSystem As IJPlateSystem: Set pPlateSystem = pGeometricConstruction.Input("PlateSystem")
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro
    Set pGeometricConstructionMacro = PlateSystem_GetMacro(pPlateSystem)
    
    Dim pElementsOfChildren As IJElements
    Set pElementsOfChildren = Entity_GetElementsOfRelatedEntities(pPlateSystem, "IJDesignParent", "Components")
    
    Dim bFound As Boolean: Let bFound = False
    Dim i As Integer
    For i = 1 To pElementsOfChildren.Count
        If TypeOf pElementsOfChildren(i) Is IJGeometricConstructionMacro Then
            Let bFound = True
            Exit For
        End If
    Next
    
    If Not bFound Then
'''        Call Entity_ConnectEntity(pPlateSystem, pGeometricConstructionMacro, "IJDesignParent", "Components")
    End If
End Sub
'''Function PlateSystem_GetMacro(pPlateSystem As IJPlateSystem) As IJGeometricConstructionMacro
'''    Dim pStructCustomGeometry As IJDStructCustomGeometry
'''    Set pStructCustomGeometry = pPlateSystem
'''
'''    Dim sProgid As String
'''    Dim pElementsOfParents As IJElements: Set pElementsOfParents = New JObjectCollection
'''    Call pStructCustomGeometry.GetCustomGeometry(sProgid, pElementsOfParents)
'''
'''    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro
'''    Set pGeometricConstructionMacro = Entity_GetElementsOfRelatedEntities(pElementsOfParents(1), "IJGeometry", "ConstructionForOutput").Item(1)
'''
'''    Set PlateSystem_GetMacro = pGeometricConstructionMacro
'''End Function
'''Public Function Entity_GetElementsOfRelatedEntities(oSource As Object, sNameOfRelation As String, sNameOfRole As String) As IJElements
'''    Dim pAssocRelations As IJDAssocRelation
'''    Set pAssocRelations = oSource
'''
'''    Dim pRelationshipCol As IJDRelationshipCol
'''    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfRelation, sNameOfRole)
'''
'''    Dim pElementsOfRelatedEntities As IJElements
'''    Set pElementsOfRelatedEntities = New JObjectCollection
'''    Dim lIndex As Long
'''    For lIndex = 1 To pRelationshipCol.Count
'''        Dim pRelationship As IJDRelationship
'''        Set pRelationship = pRelationshipCol.Item(lIndex)
'''
'''        Dim oTarget As Object
'''        Set oTarget = pRelationship.Target
'''
'''        Call pElementsOfRelatedEntities.Add(pRelationship.Target)
'''    Next
'''    Set Entity_GetElementsOfRelatedEntities = pElementsOfRelatedEntities
'''End Function
'''Public Sub Entity_ConnectEntity(oSource As Object, oTarget As Object, sNameOfInterface As String, sNameOfRole As String)
'''    Dim pAssocRelations As IJDAssocRelation
'''    Set pAssocRelations = oSource
'''
'''    Dim pRelationshipCol As IJDRelationshipCol
'''    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfInterface, sNameOfRole)
'''
'''    Dim pRelationShipHelper As DRelationshipHelper
'''    Call pRelationshipCol.Add(oTarget, "", pRelationShipHelper)
'''End Sub




