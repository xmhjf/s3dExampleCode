VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnipedFlangeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : SnipedFlangeParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    01/Jan/2012 - Created
'    07/Feb/2012 - svsmylav
'       TR-CP-209564: Corrected FlangeClearance from 15mm to 0.06 mm.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "SnipedFlangeParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJEndCutDrawingType
Implements IJDUserSymbolServices

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
    
    pOH.SetOutput "FlangeClearance"
    pOH.SetOutput "Slope"
    pOH.SetOutput "FlangeWidth"
    
    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
    
    pOH.SetOutput "BottomFlange"
  
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
    
    ' -----------------
    ' Get the item name
    ' -----------------
    Dim itemName As String
    itemName = pPRL.SmartOccurrence.Item
    
    sMsg = "Defining ParameterRule Outputs"
    
    ' ------------------------
    ' Set the flange clearance
    ' ------------------------
    If itemName = "StraightSetbackBoundStd_LR" Then
        pPRL.Add "FlangeClearance", 0.025
    Else
        pPRL.Add "FlangeClearance", 0.0001 '0.1 mm
    End If
    
    ' --------------------
    ' Set the flange width
    ' --------------------
    If itemName = "StraightSetbackBoundStd_LR" Then
        Dim oPortBounded As IJPort
        Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)

        Dim dFlangeWidth As Double
        
        If TypeOf oPortBounded.Connectable Is ISPSMemberPartCommon Then
            Dim oBoundedMemberPart As StructDetailObjects.MemberPart
            Set oBoundedMemberPart = New MemberPart
            Set oBoundedMemberPart.object = oPortBounded.Connectable
            dFlangeWidth = oBoundedMemberPart.FlangeLength
        ElseIf TypeOf oPortBounded.Connectable Is IJProfile Then
            Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
            Set oBoundedProfilePart = New ProfilePart
            Set oBoundedProfilePart.object = oPortBounded.Connectable
            dFlangeWidth = oBoundedProfilePart.FlangeLength
        End If

        If dFlangeWidth <= 0.15 Then
            pPRL.Add "FlangeWidth", 0.025
        Else
            pPRL.Add "FlangeWidth", 0.05
        End If
    Else
        pPRL.Add "FlangeWidth", 0.05  '50 mm
    End If
    
    ' -------------
    ' Set the slope
    ' -------------
    Dim dAngle As Double
    Dim dPI As Double
    dPI = Atn(1) * 4#
    
    If itemName = "StraightSetbackBoundStd_LR" Then
        dAngle = 45 * dPI / 180
    Else
        dAngle = 60 * dPI / 180
    End If
    
    pPRL.Add "Slope", dAngle

    ' -----------------------
    ' Set 'BottomFlange' flag
    ' -----------------------
    Dim sBottomFlange As String
    Dim bIsBottomFlange As Boolean
    
    GetSelectorAnswer pPRL, "BottomFlange", sBottomFlange

    If sBottomFlange = "No" Then
      bIsBottomFlange = False
        pPRL.Add "BottomFlange", 0
    Else
      bIsBottomFlange = True
        pPRL.Add "BottomFlange", 1
    End If

    ' -----------------
    ' Set cutting depth
    ' -----------------
    Dim dCuttingDepth As Double
    dCuttingDepth = GetAppropriateCutDepth(pPRL, bIsBottomFlange)
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
    Dim dAngle As Double
    Dim dFlangeClearance As Double
   
    dAngle = GetCustomAttribute(pSmartOccurrence, "IJUAMbrFreeEndCuts", "Slope")
    dFlangeClearance = GetCustomAttribute(pSmartOccurrence, "IJUAMbrEndCut", "FlangeClearance")
    
    If IsValueBetween(dAngle, 0, degreeToRadian(89.99)) Or IsValueBetween(dAngle, degreeToRadian(90.01), degreeToRadian(180)) Then
        
        If GreaterThan(dFlangeClearance, 0.004) Then
            IJEndCutDrawingType_UpdateDrawingType = Sniped_Offset_FlangeCuts
        Else
            IJEndCutDrawingType_UpdateDrawingType = Sniped_No_Offset_FlangeCuts
        End If
    Else
        If GreaterThan(dFlangeClearance, 0.004) Then
            IJEndCutDrawingType_UpdateDrawingType = Straight_Offset_FlangeCuts
        Else
            IJEndCutDrawingType_UpdateDrawingType = Straight_No_Offset_FlangeCuts
        End If
    End If
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, IJEndCutDrawingType_UpdateDrawingType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
    
End Function

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
