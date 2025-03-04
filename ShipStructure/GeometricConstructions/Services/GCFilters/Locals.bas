Attribute VB_Name = "Locals"
Option Explicit
Public Function PlateSystem_GetMacro(pPlateSystem As IJPlateSystem) As IJGeometricConstructionMacro
    If TypeOf pPlateSystem Is IJDStructCustomGeometry Then
        Dim pStructCustomGeometry As IJDStructCustomGeometry
        Set pStructCustomGeometry = pPlateSystem
        
        Dim sProgid As String
        Dim pElementsOfParents As IJElements: Set pElementsOfParents = New JObjectCollection
        On Error Resume Next
        Call pStructCustomGeometry.GetCustomGeometry(sProgid, pElementsOfParents)
        On Error GoTo 0
        
        If Err.Number = 0 And Not pElementsOfParents Is Nothing Then
            If pElementsOfParents.Count > 0 Then
                Dim pElementsOfGrandParents As IJElements
                Set pElementsOfGrandParents = Entity_GetElementsOfRelatedEntities(pElementsOfParents(1), "IJGeometry", "ConstructionForOutput")
                If Not pElementsOfGrandParents Is Nothing Then
                    If pElementsOfGrandParents.Count > 0 Then
                        Dim oGeometricConstructionMacro As Object
                        Set oGeometricConstructionMacro = pElementsOfGrandParents.Item(1)
                          
                        If TypeOf oGeometricConstructionMacro Is IJGeometricConstructionMacro Then
                            Set PlateSystem_GetMacro = oGeometricConstructionMacro
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
Function Entity_GetElementsOfRelatedEntities(oSource As Object, sNameOfInterface As String, sNameOfRole As String) As IJElements
    Dim pAssocRelations As IJDAssocRelation
    Set pAssocRelations = oSource
    
    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfInterface, sNameOfRole)
    
    Dim pElementsOfRelatedEntities As IJElements
    Set pElementsOfRelatedEntities = New JObjectCollection
    Dim lIndex As Long
    For lIndex = 1 To pRelationshipCol.Count
        Dim pRelationship As IJDRelationship
        Set pRelationship = pRelationshipCol.Item(lIndex)
        
        Dim oTarget As Object
        Set oTarget = pRelationship.Target
    
        Call pElementsOfRelatedEntities.Add(pRelationship.Target)
    Next
    Set Entity_GetElementsOfRelatedEntities = pElementsOfRelatedEntities
End Function



