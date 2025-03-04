Attribute VB_Name = "ServiceMgrRuleHelper"
'********************************************************************
' Copyright (C) 2011 Intergraph Corporation.  All Rights Reserved.
'
' File: ServiceMgrRuleHelper.bas
'
' Author: Siva
'
' Abstract: Mfg Service Manager rule helper methods
'
'  History:
'   20th JUNE 2011  Siva     Creation
'********************************************************************

Option Explicit

Const MODULE = "ServiceMgrRuleHelper.bas"
Private Const IID_IJCommonPartMember As String = "{0E8CCDC9-C434-46C6-8413-036A767D8462}"

Public Function GetProductionRoutingObject(ByVal oPart As Object) As Object
    Const METHOD = "GetProductionRoutingObject"
    On Error GoTo ErrorHandler
    
    Dim oPlProdRoutService As PlanningObjects.PlnProdRouting
    Set oPlProdRoutService = New PlanningObjects.PlnProdRouting
    
    Set oPlProdRoutService.object = oPart
    
    ' Get the Production routing object
    Dim oProdRouting    As IJDProductionRouting
    Set oProdRouting = oPlProdRoutService.GetProductionRouting
    
    Set GetProductionRoutingObject = oProdRouting
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub GetParentPartAndAssembly(oMfgObject As Object, oParentPart As Object, oParentAssembly As Object)
Const METHOD = "GetParentPartAndAssembly"
On Error GoTo ErrorHandler
    
    ' Get the parent part of MfgObject
    Dim oMfgChild As IJMfgChild
    Set oMfgChild = oMfgObject
    Set oParentPart = oMfgChild.GetParent
    
    ' Note that for Marking Lines, the parent is MarkingFolder
    If TypeOf oMfgObject Is IJMfgMarkingLines_AE Then
        Set oMfgChild = oParentPart
        Set oParentPart = oMfgChild.GetParent
    End If
    
    ' Get the parent assembly
    If Not oParentPart Is Nothing Then
        Set oMfgChild = Nothing
        Set oMfgChild = oParentPart
        
        Dim oObject As Object
        Set oObject = oMfgChild.GetParent
        
        If TypeOf oObject Is IJAssemblyBase Then
            Set oParentAssembly = oObject
        Else
            Dim oAssemblyChild As IJAssemblyChild
            Set oAssemblyChild = oObject
            Set oParentAssembly = oAssemblyChild.Parent
        End If
    End If
    
    Exit Sub
    
Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function GetCommonPartGroup(oPart As Object) As Object
Const METHOD = "GetCommonPartGroup"
On Error GoTo ErrorHandler

    Dim oObject         As Object
    Dim oAssocRel       As IJDAssocRelation
    Dim oTargetObjCol   As IJDTargetObjectCol
    
    Set oAssocRel = oPart

    If oAssocRel Is Nothing Then
        Exit Function
    End If

    Set oTargetObjCol = oAssocRel.CollectionRelations(IID_IJCommonPartMember, "CommonParent")

    If Not oTargetObjCol Is Nothing Then
        If oTargetObjCol.Count = 1 Then
            
            If TypeOf oTargetObjCol.Item(1) Is IJCommonPartGroup Then
                GetCommonPartGroup = oTargetObjCol.Item(1)
            End If
            
        End If
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetKnucklesOnProfilePart(oPart As Object) As IJElements
Const METHOD = "GetKnucklesOnProfilePart"
On Error GoTo ErrorHandler

    Dim oSemanticsUtil  As IJProfilePartSemanticsUtil
    Set oSemanticsUtil = New ProfilePartSemanticsUtil
    
    Dim oStructDetailHelper As IJStructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    
    Dim oLeafSystem    As Object
    oStructDetailHelper.IsPartDerivedFromSystem oPart, oLeafSystem
    
    Dim oProfileKnuckles    As IJElements
    Set oProfileKnuckles = New JObjectCollection
    
    Dim oTempProfKnuckles    As IJElements
    oSemanticsUtil.GetProfileKnucklesForPart oLeafSystem, pkmmBend, oTempProfKnuckles
    oProfileKnuckles.AddElements oTempProfKnuckles
    Set oTempProfKnuckles = Nothing
    
    oSemanticsUtil.GetProfileKnucklesForPart oLeafSystem, pkmmIgnore, oTempProfKnuckles
    oProfileKnuckles.AddElements oTempProfKnuckles
    Set oTempProfKnuckles = Nothing
    
    oSemanticsUtil.GetProfileKnucklesForPart oLeafSystem, pkmmSplit, oTempProfKnuckles
    oProfileKnuckles.AddElements oTempProfKnuckles
    Set oTempProfKnuckles = Nothing
    
    Set GetKnucklesOnProfilePart = oProfileKnuckles
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetOpeningsOnPart(oPart As Object) As IJElements
Const METHOD = "GetOpeningsOnPart"
On Error GoTo ErrorHandler

    ' Get Openings from plate part
    Dim oSystem As Object
    Dim oChild As IJSystemChild
    Set oChild = oPart
    Dim oUnkPart As Object
    Dim oTempPlate As IJPlate
    Dim oTempProfile As IJProfile
    Dim OpID As StructOperation
    
    ' Get plate/profile system recursively. If plate/profile System is split, get the parent.
    Do
        Set oUnkPart = Nothing
        Set oUnkPart = oChild.GetParent
        
        If Not TypeOf oUnkPart Is IJPlate Or Not TypeOf oUnkPart Is IJProfile Then
            Exit Function    ' Parent does not have to be a plate/profile and that is OK
        End If
        
        OpID = NoOperation ' Initialize
        If TypeOf oPart Is IJPlatePart Then
            Set oTempPlate = oUnkPart
            OpID = oTempPlate.OperationHistory
        Else
            Set oTempProfile = oUnkPart
            OpID = oTempProfile.OperationHistory
        End If
        
        Set oChild = oTempPlate
    Loop While (OpID & SplitOperation)
    
    Set oSystem = oUnkPart
    
    ' Get openings from plate/profile system.
    Dim oPartInfo   As IJDPartInfo
    Set oPartInfo = New PartInfo
    
    Dim oOpeningElements    As IJElements
    Set oOpeningElements = oPartInfo.PlateOpenings(oSystem)

    Set GetOpeningsOnPart = oOpeningElements
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

