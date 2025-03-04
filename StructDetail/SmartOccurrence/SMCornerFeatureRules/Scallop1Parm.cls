VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scallop1Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : Snipe1Parm.cls
'
'  Description :
'       It computes the item paramaters in the context of the smart occurrence.
'
'       - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'       - Outputs defined by name the collection of parameters
'
'  Author      : Alligators
'
'  History     :
'    18/APR/2011 - Creation
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
'*********************************************************************************************
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature
Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.Scallop1Parm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.Scallop1Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\Scallop1Parm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
    
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
    
    pOH.SetOutput "Radius"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler

    'Get corner feature part object
    Dim oCF As IJSDOCornerFeature
    Set oCF = New StructDetailObjectsex.CornerFeature
    Set oCF.object = pPRL.SmartOccurrence

    Dim oPart As Object
    Set oPart = oCF.GetPartObject
    
    Dim oCFChild As IJDesignChild
    Dim oParentSlot As Object

    Set oCFChild = oCF.object
    Set oParentSlot = oCFChild.GetParent
    
    Dim oSDFeatureUtils As IJSDFeatureAttributes
    Set oSDFeatureUtils = New SDFeatureUtils
    Dim oFacePort As Object, oEdgePort1 As Object, oEdgePort2 As Object
    
    oSDFeatureUtils.get_CornerCutInputsEx pPRL.SmartOccurrence, oFacePort, oEdgePort1, oEdgePort2
    
    If TypeOf oParentSlot Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oParentSlot
        'Feature on slot
        If oFeature.get_StructFeatureType = SF_Slot Then
            
            Dim bIsBaseCF As Boolean, bIsInsideCF As Boolean, bIsOuterCF As Boolean
            DetermineCFPositionOnSlot oEdgePort1, oEdgePort2, bIsBaseCF, bIsInsideCF, bIsOuterCF
            
            If bIsBaseCF Then
                Dim oSlotWrapper As New StructDetailObjects.Slot
                Set oSlotWrapper.object = oFeature
                
                Dim dSectionHeight As Double
                If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
                    Dim oProfilePenetrating As New StructDetailObjects.ProfilePart
                    Set oProfilePenetrating.object = oSlotWrapper.Penetrating
                    
                    dSectionHeight = oProfilePenetrating.Height
                ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
                    Dim oPlatePenetrating As New StructDetailObjects.PlatePart
                    Set oPlatePenetrating.object = oSlotWrapper.Penetrating
                    
                    Dim oSlotMappingRule As IJSlotMappingRule
                    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
                    
                    oSlotMappingRule.GetSectionDepth oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dSectionHeight
                Else
                    'Unsupported Penetrating Section
                    dSectionHeight = 0
                End If

                If dSectionHeight <= 0.2 Then
                    pPRL.Add "Radius", 0.03
                ElseIf dSectionHeight > 0.2 And dSectionHeight <= 0.3 Then
                    pPRL.Add "Radius", 0.05
                ElseIf dSectionHeight > 0.3 Then
                    pPRL.Add "Radius", 0.075
                End If

                Exit Sub
            ElseIf bIsInsideCF Then
                pPRL.Add "Radius", 0.015
                Exit Sub
            End If
        End If
    End If

    Set oFacePort = Nothing
    Set oEdgePort1 = Nothing
    Set oEdgePort2 = Nothing
    Set oParentSlot = Nothing
    Set oSDFeatureUtils = Nothing
    Set oCFChild = Nothing
    
    Dim oAssemConn As StructDetailObjects.AssemblyConn
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    Dim dShorter As Double
    Dim dLength1 As Double
    Dim dLength2 As Double
    Dim oAC As Object
    
    AssemblyConnection_SmartItemName oCF.object, , oAC, 4
    
    Dim eOACType As eACType
    
    'Get AC type
    If Not oAC Is Nothing Then
         eOACType = GetMbrAssemblyConnectionType(oAC)
    End If

    ' Parameter rule:  Set scallop radius
    If TypeOf oPart Is IJPlate Then
        pPRL.Add "Radius", 0.05
        
    ' CF on Stiffener
    ElseIf TypeOf oPart Is IJStiffener Then     'oPart is a profile
    
        If Not eOACType = ACType_None Then
            Set oConnAttrbs = New SDConnectionUtils
            Set oAssemConn = New StructDetailObjects.AssemblyConn
            Set oAssemConn.object = oAC
        
            If GreaterThanOrEqualTo(oConnAttrbs.get_AuxiliaryPorts(oAC).Count, 1) Or _
                                oAssemConn.ConnectedObject2Type = SDOBJECT_MEMBER Then
                
                Set oConnAttrbs = Nothing
                Set oAssemConn = Nothing
                'Get CF Edge on Bounding Lengths
                GetCornerFatureEdgeLengths oCF.object, oAC, pPRL.InputObject(INPUT_PORT2EDGE), _
                                       pPRL.InputObject(INPUT_PORT3EDGE), dLength1, dLength2
                Set oCF = Nothing
                
                'If Lengths are zero then set Radius by Profile Height
                If dLength1 = 0 Or dLength2 = 0 Then
                    'Set CF Radius Based on Profile height
                Else
                    'Get the Shorter Length of CF Edges
                    If LessThanOrEqualTo(dLength1, dLength2) Then
                        dShorter = dLength1
                    Else
                        dShorter = dLength2
                    End If
                    'Set Snipe Radius
                    Select Case dShorter
                        Case 0 To 0.2
                            pPRL.Add "Radius", 0.015
                        Case 0.2 To 0.28
                            pPRL.Add "Radius", 0.025
                        Case 0.28 To 0.4
                            pPRL.Add "Radius", 0.035
                        Case Is >= 0.4
                            pPRL.Add "Radius", 0.05
                    End Select
                    Set oPart = Nothing
                    Exit Sub
                End If
            End If
          End If
    
        Dim oProfile As StructDetailObjects.ProfilePart
        Set oProfile = New StructDetailObjects.ProfilePart
        Set oProfile.object = oPart
        
        'Set CF Radius by Profile Height
        Select Case oProfile.Height
            Case 0 To 0.2
                pPRL.Add "Radius", 0.035
            Case 0.2 To 0.4
                pPRL.Add "Radius", 0.05
            Case Is >= 0.4
                pPRL.Add "Radius", 0.075
        End Select
        
        Set oProfile = Nothing
        Set oPart = Nothing
    ElseIf TypeOf oPart Is IJBeam Then
        Dim oBeam As New StructDetailObjects.BeamPart
        Set oBeam.object = oPart
        Select Case oBeam.Height
            Case 0 To 0.2
                pPRL.Add "Radius", 0.035
            Case 0.2 To 0.4
                pPRL.Add "Radius", 0.05
            Case Is >= 0.4
                pPRL.Add "Radius", 0.075
        End Select
        
        Set oBeam = Nothing
        Set oPart = Nothing
        
    ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
    
        If Not eOACType = ACType_None Then
            'Get CF Edge on Bounding Lengths
            GetCornerFatureEdgeLengths oCF.object, oAC, pPRL.InputObject(INPUT_PORT2EDGE), _
                                   pPRL.InputObject(INPUT_PORT3EDGE), dLength1, dLength2
            Set oCF = Nothing
            
            If dLength1 = 0 Or dLength2 = 0 Then
                'Set Parameter values based on Height of the Member
            Else
                'Get the Shorter Length of CF Edges
                If LessThanOrEqualTo(dLength1, dLength2) Then
                    dShorter = dLength1
                Else
                    dShorter = dLength2
                End If
                'Parameter Values Based on CF Edges
                Select Case dShorter
                    Case 0 To 0.2
                        pPRL.Add "Radius", 0.015
                    Case 0.2 To 0.28
                        pPRL.Add "Radius", 0.025
                    Case 0.28 To 0.4
                        pPRL.Add "Radius", 0.035
                    Case Is >= 0.4
                        pPRL.Add "Radius", 0.05
                End Select
                Set oPart = Nothing
                Exit Sub
            End If
        End If
        
        'Set CF Radius
        Dim oMember As New StructDetailObjects.MemberPart
        Set oMember.object = oPart
        Select Case oMember.Height
            Case 0 To 0.2
                pPRL.Add "Radius", 0.035
            Case 0.2 To 0.4
                pPRL.Add "Radius", 0.05
            Case Is >= 0.4
                pPRL.Add "Radius", 0.075
        End Select
        Set oMember = Nothing
        Set oPart = Nothing
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
 
End Sub
  
' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
' ** End CM **
'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
  
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
    On Error Resume Next
    
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
    pPR.IJDInputs.RemoveAllInput
    pPR.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractParameterRule pPR
    
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pPR
    pIH.InitAs m_FamilyProgid
    ParameterRuleInputs pIH
    
    Dim pOH As IJDOutputsHelper
    Set pOH = New OutputHelper
    pOH.Representation = pPR.IJDRepresentations.Item(1)
    pOH.InitAs m_FamilyProgid
    ParameterRuleOutputs pOH
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
    Dim pPRL As IJDParameterLogic
    Set pPRL = New ParameterLogic
    pPRL.Representation = pRep
    ParameterRuleLogic pPRL
  
End Sub

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

