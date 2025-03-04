VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snipe1Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-16, Intergraph Corporation.  All rights reserved.
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
'
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
'   07 Jan 2016 -GHM       - TR-279987- Updated GetCornerFatureEdgeLengths() method with proper corner feature input object
'*********************************************************************************************

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.Snipe1Parm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.Snipe1Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\Snipe1Parm.cls"

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
    
    pOH.SetOutput "Ulength"
    pOH.SetOutput "Vlength"
    pOH.SetOutput "Flip"
    
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
    Dim oCornerFeature As IJSDOCornerFeature
    Set oCornerFeature = New StructDetailObjectsex.CornerFeature
    Set oCornerFeature.object = pPRL.SmartOccurrence
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    
    Dim oPart As Object
    Set oPart = oCornerFeature.GetPartObject
        
    ' Get Corner Feature Orientation Answer
    Dim strCornerFlip As String
    
    GetSelectorAnswer pPRL, gsCornerFlip, strCornerFlip
    
    ' Parameter rule: Set corner feature input orientation based on Answer
    ' IJDParameterlogic calls FeatureSymbol.ASCornerFeature, where:
    '    0 = NoFlip
    '    1 = Flip
    Select Case strCornerFlip
        Case "Flip"
            pPRL.Add "Flip", 1
        Case Else
            pPRL.Add "Flip", 0
    End Select
    
    Dim oSDFeatureUtils As IJSDFeatureAttributes
    Set oSDFeatureUtils = New SDFeatureUtils
    Dim oFacePort As Object, oEdgePort1 As Object, oEdgePort2 As Object
    
    oSDFeatureUtils.get_CornerCutInputsEx pPRL.SmartOccurrence, oFacePort, oEdgePort1, oEdgePort2
    
    Dim oCFChild As IJDesignChild
    Dim oParentSlot As Object

    Set oCFChild = oCornerFeature.object
    Set oParentSlot = oCFChild.GetParent

    If TypeOf oParentSlot Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oParentSlot
        'feature on slot
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
                    
                    Dim oSlotmappingRule As IJSlotMappingRule
                    Set oSlotmappingRule = CreateSlotMappingRuleSymbolInstance
                    
                    oSlotmappingRule.GetSectionDepth oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dSectionHeight
                Else
                    'Unsupported Penetrating Section
                    dSectionHeight = 0
                End If
                
                If dSectionHeight <= 0.2 Then
                    pPRL.Add "Ulength", 0.03
                    pPRL.Add "Vlength", 0.03
                ElseIf dSectionHeight > 0.2 And dSectionHeight <= 0.3 Then
                    pPRL.Add "Ulength", 0.05
                    pPRL.Add "Vlength", 0.05
                ElseIf dSectionHeight > 0.3 Then
                    pPRL.Add "Ulength", 0.075
                    pPRL.Add "Vlength", 0.075
                End If
                
                Exit Sub
            ElseIf bIsInsideCF Then
                pPRL.Add "Ulength", 0.015
                pPRL.Add "Vlength", 0.015
            Exit Sub
            End If
        End If
    End If
    
    Dim oAssemConn As StructDetailObjects.AssemblyConn
    Dim dProfileHeight As Double
    Dim dShorter As Double
    Dim dLength1 As Double
    Dim dLength2 As Double
    Dim oAC As Object
    
    AssemblyConnection_SmartItemName oCornerFeature.object, , oAC, 4
    
    Dim eOACType As eACType
    'Get AC type
    If Not oAC Is Nothing Then
         eOACType = GetMbrAssemblyConnectionType(oAC)
    End If
        
    If TypeOf oPart Is IJStiffener Then
        'oPart is a profile or edge reinforcement
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oPart
        dProfileHeight = oProfile.Height
        
        'Get Assembly Connection
        Dim strParentname As String
        Dim oAppConn As Object
        Dim oBounding As StructDetailObjects.ProfilePart
        Dim dBoundingProfileWdth As Double
          
        Parent_SmartItemName oCornerFeature.object, strParentname, oAppConn
        
        If Not oAppConn Is Nothing Then
            If TypeOf oAppConn Is IJAssemblyConnection Then
                Set oConnAttrbs = New SDConnectionUtils
                Set oAssemConn = New StructDetailObjects.AssemblyConn
                Set oAssemConn.object = oAppConn
                If oAssemConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then
                    Set oBounding = New StructDetailObjects.ProfilePart
                    If (oConnAttrbs.get_AuxiliaryPorts(oAppConn).Count >= 1) Then
                        Set oBounding.object = oAssemConn.ConnectedObject2
                        If oBounding.sectionType = "FB" Then
                             dBoundingProfileWdth = oBounding.webThickness
                        Else
                             dBoundingProfileWdth = oBounding.Width
                        End If
                        
                        Dim oAuxPort As IJPort
                        Set oAuxPort = oConnAttrbs.get_AuxiliaryPorts(oAppConn).Item(1)
                        ' For only auxilary ports
                        If TypeOf oAuxPort.Connectable Is IJPlate Then
                            Dim oAuxPlate As StructDetailObjects.PlatePart
                            Dim dAuxPlate As Double
                            
                            Set oAuxPlate = New StructDetailObjects.PlatePart
                            Set oAuxPlate.object = oAuxPort.Connectable
                            dAuxPlate = oAuxPlate.PlateThickness
                            
                            If GreaterThan(dBoundingProfileWdth - dAuxPlate, 0.02, 0) Then
                                pPRL.Add "Ulength", (dBoundingProfileWdth - dAuxPlate) - 0.02
                                pPRL.Add "Vlength", (dBoundingProfileWdth - dAuxPlate) - 0.02
                            ElseIf GreaterThan(dBoundingProfileWdth - dAuxPlate, 0.01, 0) Then
                                pPRL.Add "Ulength", 0.01
                                pPRL.Add "Vlength", 0.01
                            Else
                                pPRL.Add "Ulength", 0.005
                                pPRL.Add "Vlength", 0.005
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        If Not eOACType = ACType_None Then
            Set oConnAttrbs = New SDConnectionUtils
            
            Set oAssemConn = New StructDetailObjects.AssemblyConn
            Set oAssemConn.object = oAC
        
            If GreaterThanOrEqualTo(oConnAttrbs.get_AuxiliaryPorts(oAC).Count, 1) Or _
                        oAssemConn.ConnectedObject2Type = SDOBJECT_MEMBER Then
                                
                'Get CF Edge on Bounding Lengths
                GetCornerFatureEdgeLengths pPRL.SmartOccurrence, oAC, pPRL.InputObject(INPUT_PORT2EDGE), _
                                       pPRL.InputObject(INPUT_PORT3EDGE), dLength1, dLength2
                
                If dLength1 = 0 Or dLength2 = 0 Then
                    'Set Parameter values based on Height
                Else
             
                    'Get the Shorter Length of CF Edges
                    If LessThanOrEqualTo(dLength1, dLength2) Then
                        dShorter = dLength1
                    Else
                        dShorter = dLength2
                    End If
                    
                'Set Parameter Values
                Select Case dShorter
                    Case Is <= 0.2
                        pPRL.Add "Ulength", 0.015
                        pPRL.Add "Vlength", 0.015
                        
                    Case 0.2 To 0.28
                        pPRL.Add "Ulength", 0.025
                        pPRL.Add "Vlength", 0.025
                        
                    Case 0.28 To 0.4
                        pPRL.Add "Ulength", 0.035
                        pPRL.Add "Vlength", 0.035
                        
                    Case Is > 0.4
                        pPRL.Add "Ulength", 0.05
                        pPRL.Add "Vlength", 0.05
                End Select
                Exit Sub
                End If
            End If
          End If
        
    ElseIf TypeOf oPart Is IJBeam Then
        Dim oBeam As New StructDetailObjects.BeamPart
        Set oBeam.object = oPart
        dProfileHeight = oBeam.Height
        
    ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
        
        Dim oStructPort As IJStructPort
        Set oStructPort = oEdgePort1
        
        Dim dShorterEdgeLength As Double
        
        If oStructPort.SectionID >= e_JXSEC_MultipleBounding_5001 And oStructPort.SectionID <= e_JXSEC_MultipleBounding_5005 Then

            GetCornerFatureEdgeLengths pPRL.SmartOccurrence, Nothing, oEdgePort1, oEdgePort2, dLength1, dLength2, oFacePort
        
            If dLength1 < dLength2 Then
                dShorterEdgeLength = dLength1
            Else
                dShorterEdgeLength = dLength2
            End If
            
            Dim dShoterLenghthinmm As Double
            dShoterLenghthinmm = dShorterEdgeLength * 1000
            
            'Set Parameter Values
            Select Case dShoterLenghthinmm

                Case Is <= 20 'mm
                    pPRL.Add "Ulength", dShorterEdgeLength / 2
                    pPRL.Add "Vlength", dShorterEdgeLength / 2
                Case 20 To 200 'mm
                    pPRL.Add "Ulength", 0.015
                    pPRL.Add "Vlength", 0.015

                Case 200 To 280 'mm
                    pPRL.Add "Ulength", 0.025
                    pPRL.Add "Vlength", 0.025

                Case 280 To 400 'mm
                    pPRL.Add "Ulength", 0.035
                    pPRL.Add "Vlength", 0.035

                Case Is > 400 'mm
                    pPRL.Add "Ulength", 0.05
                    pPRL.Add "Vlength", 0.05
            End Select
        Exit Sub
        End If
        
        If Not eOACType = ACType_None Then
            'Get CF Edge on Bounding Lengths
            GetCornerFatureEdgeLengths oCornerFeature.object, oAC, pPRL.InputObject(INPUT_PORT2EDGE), _
                                   pPRL.InputObject(INPUT_PORT3EDGE), dLength1, dLength2
            Set oCornerFeature = Nothing
            Set oPart = Nothing
            
            If dLength1 = 0 Or dLength2 = 0 Then
                pPRL.Add "Ulength", 0.015
                pPRL.Add "Vlength", 0.015
                Exit Sub
            Else
            
                'Get the Shorter Length of CF Edges
                If LessThanOrEqualTo(dLength1, dLength2) Then
                    dShorter = dLength1
                Else
                    dShorter = dLength2
                End If
                'Set Parameter Values
                Select Case dShorter
                    Case Is <= 0.2
                        pPRL.Add "Ulength", 0.015
                        pPRL.Add "Vlength", 0.015
                        
                    Case 0.2 To 0.28
                        pPRL.Add "Ulength", 0.025
                        pPRL.Add "Vlength", 0.025
                        
                    Case 0.28 To 0.4
                        pPRL.Add "Ulength", 0.035
                        pPRL.Add "Vlength", 0.035
                        
                    Case Is > 0.4
                        pPRL.Add "Ulength", 0.05
                        pPRL.Add "Vlength", 0.05
                End Select
                Exit Sub
            End If
        End If
                
        pPRL.Add "Ulength", 0.015
        pPRL.Add "Vlength", 0.015
        
    Else
        ' oPart is a plate or an unknown type.  Default to 50mm
        If TypeOf oPart Is IJPlate Then
            pPRL.Add "Ulength", 0.05
            pPRL.Add "Vlength", 0.05
            Exit Sub
        End If
    End If

    If Not TypeOf oPart Is ISPSMemberPartPrismatic Then
        Select Case dProfileHeight
            Case 0 To 0.2
                pPRL.Add "Ulength", 0.035
                pPRL.Add "Vlength", 0.035
            Case 0.2 To 0.4
                pPRL.Add "Ulength", 0.05
                pPRL.Add "Vlength", 0.05
            Case Is >= 0.4
                pPRL.Add "Ulength", 0.075
                pPRL.Add "Vlength", 0.075
        End Select
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
      
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
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

