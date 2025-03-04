VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlgInHookFlgRlfSnpTopWLParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : FlgInHookFlgRlfSnpTopWLParm.cls
'
'  Description :
'       Sets parameters for the FlgInHookFlgRlfSnpTopWL cut.
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "FlgInHookFlgRlfSnpTopWLParm"
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
    
    pOH.SetOutput "Radius"
    pOH.SetOutput "TopEdgeOffset"
    pOH.SetOutput "TopOffset"
    pOH.SetOutput "TopEdgeInsideOffset"
    pOH.SetOutput "Slope"
    pOH.SetOutput "SlopeOutside"
    pOH.SetOutput "WebThickness"
    pOH.SetOutput "WebClearance"
    pOH.SetOutput "BottomLength"
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
Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
    
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    Dim dEdgeOffset As Double
    Dim dRadius As Double
    Dim dEdgeSetback As Double
    Dim dEdgeInsideOffset As Double
    Dim dFlangeOffset As Double ' Only for symbols with a left flange.  For this symbol we want the web thickness.
    Dim dCutDepth As Double
    
    Dim oParmCM As New EndCutParmCM
    oParmCM.ComputeHookFlangReliefParameters oPRL, _
                                             False, _
                                             False, _
                                             dEdgeOffset, _
                                             dRadius, _
                                             dEdgeSetback, _
                                             dEdgeInsideOffset, _
                                             dFlangeOffset, _
                                             dCutDepth
    
    Dim dSlope As Double
    Dim dSlopeOutside As Double ' not used
    Dim dClearance As Double ' not used
    Dim dNose As Double
    
    oParmCM.GetOutsideCutSnipeParameters oPRL, False, dSlope, dSlopeOutside, dClearance, dNose
    
    oPRL.Add "BottomLength", dEdgeOffset + dEdgeSetback - (2 * dRadius)
    oPRL.Add "Radius", dRadius
    oPRL.Add "TopEdgeOffset", dEdgeSetback
    oPRL.Add "TopEdgeInsideOffset", dEdgeInsideOffset
    oPRL.Add "CutDepth", dCutDepth
    oPRL.Add "TopOffset", dNose
    oPRL.Add "Slope", dSlope
    oPRL.Add "SlopeOutside", dSlopeOutside
    oPRL.Add "WebThickness", oParmCM.GetWebOrFlangeThickness(oPRL, False)
    oPRL.Add "WebClearance", dClearance
    
    oParmCM.SetBottomFlangeParameter oPRL

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
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
    Dim sDrawingtype As String
    Dim dBottomLength As Double
    
    dBottomLength = GetCustomAttribute(pSmartOccurrence, "IJUAMbrEndCut", "BottomLength")
    
    If GreaterThan(dBottomLength, 0.004) Then
        sDrawingtype = Sniped_Offset_FlangeCuts
    Else
        sDrawingtype = Sniped_No_Offset_FlangeCuts
    End If
    
    
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.TopOrLeft, sDrawingtype
    IJEndCutDrawingType_UpdateDrawingType = sDrawingtype
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
    
End Function
'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************
