VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebNonPenParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenWebNonPenParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'  10/Feb/2012 - TR209316 - CM - Updated Cutting Depth of Web Cut from 0.01mm to 0.1mm
'                                 as the Web Cut has to atleast touch the Flange Notch.
'    16/may/2012 - hgunturu TR-212435: Modified the cutting depth for 'Generic_2B_01' item.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "GenWebNonPenParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dCuttingDepth As Double
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    Dim oBoundingObject As Object
    
    sMsg = "Defining ParameterRule Outputs"
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    
    
    Dim oBoundedObject As Object
    Dim sACItemName As String
    Dim oACObject As Object
    Dim oNearBoundingToWeb As Object
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    
    Set oBoundedObject = oPortBounded.Connectable
    Dim dWebThickness As Double
    
    If TypeOf oBoundedObject Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBoundedObject
        dWebThickness = oBoundedMemberPart.WebThickness
    ElseIf TypeOf oBoundedObject Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New ProfilePart
        Set oBoundedProfilePart.object = oBoundedObject
        dWebThickness = oBoundedProfilePart.WebThickness
    Else
        'Special case need to be handled
    End If
    
    AssemblyConnection_SmartItemName pPRL.SmartOccurrence, sACItemName, oACObject
     
    If IsTubularMember(oBoundedObject) Or IsRectangularMember(oBoundedObject) Then
        'For all tubular members the Cutting Dpeth should be set complete Depth
        'since it doesn't have any Flanges irrespective of Web/Flange Penetrated cases
        'Only for Flanged Cross Section we give optimum CutDepth such that
        'Flange CutDepth will do remaining Cut for Flange Penetrated cases but this check
        'is not required for Tubular Cross Section(e.g PIPE, HSSC, CS, P, R etc...)
        'Hence setting the full Cut Depth
        dCuttingDepth = EndCut_GetCutDepth(pPRL)
        
    ElseIf sACItemName = "Generic_2B_01" Then
    
        dCuttingDepth = dWebThickness + 0.0001
        Set oBoundingObject = oPortBounding.Connectable
        
        If TypeOf oBoundingObject Is IJPlatePart Then
            Set oSDO_PlatePart = New StructDetailObjects.PlatePart
            Set oSDO_PlatePart.object = oBoundingObject
            
            If oSDO_PlatePart.PlateThickness > dCuttingDepth Then
                dCuttingDepth = oSDO_PlatePart.PlateThickness + 0.0001
            End If
        End If
    Else
        'define the cutting depth as Web thickness + Small Tolearcane(0.1mm) so that
        'it comepletely cuts the Web but not complete cross section and touches the Flange Notch
        dCuttingDepth = dWebThickness + 0.0001
    
    End If
    pPRL.Add "CutDepth", dCuttingDepth

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
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
    Dim pIH As IJDInputsHelper
    Dim pOH As IJDOutputsHelper
    
    pDFact.InitAbstractParameterRule pPR
    
    Set pIH = New InputHelper
    pIH.definition = pPR
    pIH.InitAs m_FamilyProgid
    ParameterRuleInputs pIH
    
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
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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
'*****************************************************************************************************
' Method Name : IJEndCutDrawingType_UpdateDrawingType
' Inputs      : feature occurence object
' Output      : string, returns the end cut type
' Description : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*****************************************************************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    IJEndCutDrawingType_UpdateDrawingType = Straight_No_Offset_WebCuts
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, IJEndCutDrawingType_UpdateDrawingType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




