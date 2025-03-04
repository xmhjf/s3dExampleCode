VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuiltUpMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2012 Intergraph Corporation.  All Rights Reserved.
'
' File: BuiltUpMembers.cls
'
' Author: Devi Kishore Adiraju
'
' Abstract: Rule to create assemblies for Built-Up Members
'
' Description:
' This rule creates assemblies of Type, "BuiltUp Member". It identifies
' the Built-Up members which are associated to the parts (under the selected assembly).
' It then creates an Assembly of type, "BuilUp Member" for each Built-up member only
' when all the parts that make the Built-up member belong to the selected assembly.
' Finally it associates each new assembly with the corresponding Built-Up member with a relation.
'
'  History
'   Jul 11 2012      Devi Kishore Adiraju      Created.
'   Jan 17 2017      Gopi M                    DI-CP-308270 Assembly properties customization.
'********************************************************************

Option Explicit

Private Const IID_IJMfgParent As String = "{07AE45B9-D88F-4DF2-865D-FEBE67D82E32}"

Implements IAutomatAssembly


Private Function IAutomatAssembly_DoAssembliesByRule(ByVal oElements As GSCADPlnIntHelper.IJElements, ByVal e_Option As GSCADPlnIntHelper.EComboOptions) As Boolean
On Error GoTo ErrorHandler

    Dim oSubAssemblies                  As IJElements           'Sub-assemblies of one of the selected assemblies
    Dim oSubAssembly                    As IJAssemblyChild      'A sub-assembly or UnAssigned Parts folder
    Dim oSubAssyAsAssyBase              As IJAssemblyBase       'Sub-assembly as AssemblyBase
    Dim oPlatePartsInAssy               As IJElements           'Plate Parts in a sub-assembly or UnAssigned Parts folder
    Dim oUAPartsFolder                  As IJAssembly           'UnAssigned Parts folder of Sub-Assembly
    Dim oPlatePartsInUAFolder           As IJElements           'Plate parts in UnAssigned parts folder
    Dim oBUMembersRelatedToAssy         As IJElements           'BU Mumbers which have atleast one of their parts under the sub-assembly/UnAssigned Parts folder
    Dim oPlatePart                      As IJAssemblyChild
    Dim oPlnIntHelper                   As IJDPlnIntHelper
    Dim oBUMember                       As ISPSDesignedMember   'BU Member
    Dim oPartsOfBUMember                As IJElements           'All parts that make a given BU Member
    Dim i                               As Long
    Dim j                               As Long
    Dim oCAutoAssembly                  As CAutoAssembly
    Dim oAssemblyParent                 As IJAssembly           'Assembly under which a new assembly needs to created
    Dim oBUMemberAssembly               As IJAssemblyBase       'New assembly (as AssemblyBase)
    Dim oNamedItem                      As IJNamedItem
    
    
    Set oCAutoAssembly = New CAutoAssembly
    Set oPlnIntHelper = New CPlnIntHelper
    
    'Loop through the selected assemblies
    For i = 1 To oElements.Count
        
        'Gather sub-assemblies based on the option selected
        Select Case e_Option
        
            'Processing all parts, so we want all sub-assemblies including UnAssigned Parts folders
            Case e_AllParts
                Set oSubAssemblies = oPlnIntHelper.GetStoredProcAssemblyChildren(oElements.Item(i), "IJAssemblyBase")
                oSubAssemblies.Add oElements.Item(i)
                
                'Need to exclude the Built-up Member assemblies.
                For Each oSubAssyAsAssyBase In oSubAssemblies
                    If oSubAssyAsAssyBase.Type = 18 Then
                        oSubAssemblies.Remove oSubAssyAsAssyBase
                    End If
                Next
                
            'Processing only UnAssigned Parts folders. So we need to look for parts in the folders themselves and create assemblies under the parents of the folders
            Case e_UnProcParts
                Set oSubAssemblies = oPlnIntHelper.GetStoredProcAssemblyChildren(oElements.Item(i), "IJPlnUnprocessedParts")
        End Select
        
        'Loop through sub assemblies
        For j = 1 To oSubAssemblies.Count
        
            'Get plate parts in the sub-assembly
            Set oSubAssembly = oSubAssemblies.Item(j)
            Set oPlatePartsInAssy = GetPlatePartsInAssembly(oSubAssembly)
                        
            'Include plate parts from UnAssigned Parts folder also, when the option is "All Parts"
            If e_Option = e_AllParts Then
                Set oSubAssyAsAssyBase = oSubAssembly
                Set oUAPartsFolder = oSubAssyAsAssyBase.GetUnProcessedParts(oSubAssyAsAssyBase)
                Set oPlatePartsInUAFolder = GetPlatePartsInAssembly(oUAPartsFolder)
                
                oPlatePartsInAssy.AddElements oPlatePartsInUAFolder
            End If
            
            'Loop through plate parts
            For Each oPlatePart In oPlatePartsInAssy
            
                'Get BuiltUp Member associated to the plate part
                Set oBUMember = GetBuiltUpMember(oPlatePart)
                
                'If there is one
                If Not oBUMember Is Nothing Then
                
                    'Get all the parts that make the BU member
                    Set oPartsOfBUMember = GetBUMemberParts(oBUMember)
                    
                    'If we did get some parts associated to the member
                    If Not oPartsOfBUMember Is Nothing Then
                    
                        'Check whether all these parts are in the sub-assembly.
                        If AreAllPartsOfBUMemberInAssy(oPartsOfBUMember, oPlatePartsInAssy) Then
                            '''''''''''''Yes: Create an Assembly for the BU Member''''''''''''''''''''''''''''''''''''
                            If TypeOf oSubAssembly Is IJPlnUnprocessedParts Then
                                Set oAssemblyParent = oSubAssembly.Parent
                            Else
                                Set oAssemblyParent = oSubAssembly
                            End If
                            
                            Set oBUMemberAssembly = oCAutoAssembly.CreateAssembly(oAssemblyParent, oPartsOfBUMember)
                            
                            'Set Bu-Member Assembly Naming rule
                            SetNamingRule oBUMemberAssembly
                            
                            'Set assembly type
                            oBUMemberAssembly.Type = 18
                            
                            'set the assembly properties defined in the rule
                            SetAssemblyPropertiesByRule oBUMemberAssembly

                            Dim oLocalCS As IJLocalCoordinateSystem
                            Dim oXVector As IJDVector
                            Dim oYVector As IJDVector
                            Dim oZVector As IJDVector

                            'Get the CS position of the build up member
                            Set oLocalCS = oBUMember
                            Set oXVector = oLocalCS.XAxis
                            Set oYVector = oLocalCS.YAxis
                            Set oZVector = oLocalCS.ZAxis
                            
                            'Set the CS vectors to the BU Assembly
                            Set oLocalCS = Nothing
                            Set oLocalCS = oBUMemberAssembly
                            
                            oLocalCS.XAxis = oXVector
                            oLocalCS.YAxis = oYVector
                            oLocalCS.ZAxis = oZVector

                            'Set default workcenters for the assembly
                            oCAutoAssembly.SetWorkAndEquipment oBUMemberAssembly
                            
                            'Relate BUmember and Assembly
                            AddRelationship oBUMemberAssembly, oBUMember, IID_IJMfgParent, "Module"
                            
                            'Remove parts from the collection
                            oPlatePartsInAssy.RemoveElements oPartsOfBUMember
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Else
                            'No: Remove parts from the collection
                            oPlatePartsInAssy.RemoveElements oPartsOfBUMember
                        End If
                    End If
                End If
            Next
        Next
    Next

wrapup:
Exit Function
ErrorHandler:
    MsgBox Err.Description, , vbInformation, "IAutomatAssembly_DoAssembliesByRule"
    GoTo wrapup
End Function

Private Function GetBuiltUpMember(oPlatePart As IJAssemblyChild) As ISPSDesignedMember
On Error GoTo ErrorHandler

    Dim oChild          As IJDesignChild
    Dim oParent         As IJDesignParent
    
    Set GetBuiltUpMember = Nothing
    
    Set oChild = oPlatePart
    Set oParent = oChild.GetParent
    
    If oParent Is Nothing Then GoTo wrapup
    
    Set oChild = oParent
    Set oParent = oChild.GetParent
    
    If oParent Is Nothing Then GoTo wrapup
    
    Set oChild = oParent
    Set oParent = oChild.GetParent
    
    If Not oParent Is Nothing Then
        If TypeOf oParent Is ISPSDesignedMember Then
            Set GetBuiltUpMember = oParent
        End If
    End If
    
wrapup:
    Set oChild = Nothing
    Set oParent = Nothing

Exit Function
ErrorHandler:
    Err.Clear
    GoTo wrapup
End Function

Private Function GetBUMemberParts(oBUMember As IJDesignParent) As IJElements
On Error GoTo ErrorHandler
    
    Dim oMemberChildren         As IJDObjectCollection
    Dim oPlateSystem            As IJPlateSystem
    Dim oSDHelper               As IJStructDetailHelper
    Dim oConvertUtils           As IJCollectionConversions
    Dim oEnumUnk                As IEnumUnknown
    Dim oVBCollection           As Collection
    Dim lIndex                  As Long
    
    oBUMember.GetChildren oMemberChildren
    
    Set GetBUMemberParts = New JObjectCollection
    Set oSDHelper = New StructDetailHelper
    Set oConvertUtils = New CCollectionConversions
    
    For Each oPlateSystem In oMemberChildren
    
        oSDHelper.GetPartsDerivedFromSystem oPlateSystem, oEnumUnk, True, False
        oConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumUnk, oVBCollection
    
        For lIndex = 1 To oVBCollection.Count
            GetBUMemberParts.Add oVBCollection.Item(lIndex)
        Next lIndex
    Next
    
wrapup:
    Set oMemberChildren = Nothing
    Set oPlateSystem = Nothing
    Set oSDHelper = Nothing
    Set oConvertUtils = Nothing
    Set oEnumUnk = Nothing
    Set oVBCollection = Nothing

Exit Function
ErrorHandler:
    Err.Clear
    GoTo wrapup
End Function

Private Function AreAllPartsOfBUMemberInAssy(oBUMemberParts As IJElements, oAllAssyChildren As IJElements) As Boolean
On Error GoTo ErrorHandler

    Dim i As Long
    
    AreAllPartsOfBUMemberInAssy = False
    
    For i = 1 To oBUMemberParts.Count
        If oAllAssyChildren.Contains(oBUMemberParts.Item(i)) = False Then
            Exit Function
        End If
    Next
    
    AreAllPartsOfBUMemberInAssy = True

Exit Function
ErrorHandler:
    Err.Clear
End Function

Private Function GetPlatePartsInAssembly(oAssembly As IJAssembly) As IJElements
On Error GoTo ErrorHandler
    
    Set GetPlatePartsInAssembly = New JObjectCollection
    
    Dim oAssyChildren           As IJDTargetObjectCol
    Dim oAsChild                As IJAssemblyChild
    Dim i                       As Long
    
    Set oAssyChildren = oAssembly.GetChildren
    
    For i = 1 To oAssyChildren.Count
    
        Set oAsChild = oAssyChildren.Item(i)
        If TypeOf oAsChild Is IJPlatePart Then
            GetPlatePartsInAssembly.Add oAsChild
        End If
    Next


Exit Function
ErrorHandler:
    Err.Clear
End Function

Private Sub AddRelationship(oAssocRelation As IJDAssocRelation, oTargetobject As Object, strInterfaceID As String, strCollectionName As String, Optional strRelationName As String = vbNullString)
On Error GoTo ErrorHandler
    
    Dim oRevision       As REVISIONLib.IJRevision
    Dim oUnkCollection  As Object
    Dim oTargetObjCol   As IJDTargetObjectCol
    Dim oRelationship   As IJDRelationship
    
    Set oRevision = New REVISIONLib.JRevision
    Set oUnkCollection = oAssocRelation.CollectionRelations(strInterfaceID, strCollectionName)
    
    ' Get collection of target objects
    Set oTargetObjCol = oUnkCollection
                        
    ' Add relationship
    oTargetObjCol.Add oTargetobject, strRelationName, oRelationship
    oRevision.AddRelationship oRelationship
    
    Set oRevision = Nothing
    Set oUnkCollection = Nothing
    Set oTargetObjCol = Nothing
    
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub SetNamingRule(oNewAssembly As Object)
On Error GoTo ErrorHandler

    Dim oBUMAssyNameRule            As GSCADGenericNamingRulesFacelets.NameRuleHolderFacelets
    Dim oRuleHelper                 As IJDNamingRulesHelper
    Dim oRules                      As IJElements
    Dim lIndex                      As Long
    Dim oTemp                       As GSCADGenericNamingRulesFacelets.NameRuleHolderFacelets
    Dim oSemantic                   As GSCADGenNameRuleAE.IJNameRuleAE
    Dim strNameRuleEntity           As String
    Dim strNameRule                 As String

    Set oRuleHelper = New GSCADNameRuleHlpr.NamingRulesHelper
    strNameRuleEntity = "CAssembly"   'CAssembly
    strNameRule = "Built-up Member Assembly Name Rule"
    
    ' Set default naming rule
    oRuleHelper.GetEntityNamingRulesGivenName strNameRuleEntity, oRules

    For lIndex = 1 To oRules.Count
        Set oTemp = oRules(lIndex)
        If oTemp.Name = strNameRule Then
          ' Return first (default) rule of collection
            Set oBUMAssyNameRule = oRules(lIndex)
            Set oTemp = Nothing
        End If
    Next lIndex
    
    If Not oBUMAssyNameRule Is Nothing Then
      oRuleHelper.AddNamingRelations oNewAssembly, oBUMAssyNameRule, oSemantic
    End If

Exit Sub
ErrorHandler:
    Err.Clear
End Sub


'Sets the properties on the given assembly based on the AssemblyProperties rule
Private Sub SetAssemblyPropertiesByRule(oBUMemberAssembly As IJAssembly)
    On Error GoTo ErrorHandler
    Dim oAssemblyHelper As PlnAssembly
    Set oAssemblyHelper = New PlnAssembly
    
    Set oAssemblyHelper.object = oBUMemberAssembly
    oAssemblyHelper.SetAssemblyPropertiesByRule True
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

