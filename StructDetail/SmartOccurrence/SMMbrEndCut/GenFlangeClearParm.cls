VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFlangeClearParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenFlangeClearParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "GenFlangeClearParm"
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
  
    pOH.SetOutput "FlangeClearance"
  
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
Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
    
    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"

    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    ' ---------------------------------------------
    ' Determine if this is the top or bottom flange
    ' ---------------------------------------------
    Dim bIsBottomFlange As Boolean
    Dim sBottomFlange As String
    
    GetSelectorAnswer oPRL, "BottomFlange", sBottomFlange

    If sBottomFlange = "No" Then
        bIsBottomFlange = False
        oPRL.Add "BottomFlange", 0
    Else
        bIsBottomFlange = True
        oPRL.Add "BottomFlange", 1
    End If
        
    ' --------------------------------------------------------
    ' Determine if there are top and bottom "outside" web cuts
    ' --------------------------------------------------------
    'Parent members dispid 2/3 creates Top/Btm WebCut, respectively.
    ' ItemByDispid will return  Nothing if the object was not created
    Dim oParentObj As Object
    Dim sParentItemName As String
    
    Parent_SmartItemName oPRL.SmartOccurrence, sParentItemName, oParentObj
    
    Dim oMemberObjects As IJDMemberObjects
    Dim oCtrWebCut As Object
    Dim oTopWebCut As Object
    Dim oBtmWebCut As Object
    
    Dim bFlangeReliefByFlangeCut As Boolean
    bFlangeReliefByFlangeCut = False
    
    If Not oParentObj Is Nothing Then
        Set oMemberObjects = oParentObj
        Set oCtrWebCut = oMemberObjects.ItemByDispid(1)
        Set oTopWebCut = oMemberObjects.ItemByDispid(2)
        Set oBtmWebCut = oMemberObjects.ItemByDispid(3)
        
        ' ------------------------------------------------------------------
        ' Identify cases where the flange is set back without an outside cut
        ' ------------------------------------------------------------------
        Dim sBottomShape As String
        Dim sTopShape As String
        Dim sShape As String
        
        GetMemberACTopAndBottomShape oPRL.SmartOccurrence, , sBottomShape, , sTopShape
        
        If (bIsBottomFlange) And (sBottomShape = gsEdgeToOutside) And (oBtmWebCut Is Nothing) Then
            bFlangeReliefByFlangeCut = True
        ElseIf (Not bIsBottomFlange) And (sTopShape = gsEdgeToOutside) And (oTopWebCut Is Nothing) Then
            bFlangeReliefByFlangeCut = True
        End If
    End If
        
    ' ---------------------------------------------------------------------------------------------------
    ' If shape indicated the edge should be clear, but there is no outside web cut, the clearance is 25mm
    ' Otherwise, the flange cut starts at the web cut point
    ' ---------------------------------------------------------------------------------------------------
    If bFlangeReliefByFlangeCut Then
        oPRL.Add "FlangeClearance", 0.025
    Else
        oPRL.Add "FlangeClearance", 0.0001
    End If
    
    ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    Dim dCuttingDepth As Double
    
    dCuttingDepth = GetAppropriateCutDepth(oPRL, bIsBottomFlange)
    oPRL.Add "CutDepth", dCuttingDepth

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
    Dim dFlangeClearance As Double
    
    dFlangeClearance = GetCustomAttribute(pSmartOccurrence, "IJUAMbrEndCut", "FlangeClearance")
    
    If GreaterThan(dFlangeClearance, 0.004) Then
        IJEndCutDrawingType_UpdateDrawingType = Straight_Offset_FlangeCuts
    Else
        IJEndCutDrawingType_UpdateDrawingType = Straight_No_Offset_FlangeCuts
    End If

    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, IJEndCutDrawingType_UpdateDrawingType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************






