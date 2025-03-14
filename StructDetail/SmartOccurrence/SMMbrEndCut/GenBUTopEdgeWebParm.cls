VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenBUTopEdgeWebParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenBUTopEdgeWebParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    01/Dec/2011 - TR-205842- Corner Features Fails when Bounding and Bounded have Equal Flange thickness
'                  Incresed All the Offset values in ParameterRuleLogic() method
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "GenBUTopEdgeWebParm"
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
    
    pOH.SetOutput "FaceOffset"
    pOH.SetOutput "TopFlangeThickness"
    pOH.SetOutput "BottomFlangeThickness"
    pOH.SetOutput "TopOffset"
    pOH.SetOutput "BottomOffset"
    pOH.SetOutput "TopEdgeOffset"
    pOH.SetOutput "TopEdgeInsideOffset"
    pOH.SetOutput "TopEdgeOutsideOffset"
    pOH.SetOutput "TopFaceOffset"
    
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
    
    sMsg = "Defining ParameterRule Outputs"
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    
    Dim dBoundedFlgThickness As Double
    If TypeOf oPortBounded.Connectable Is ISPSMemberPartCommon Then
        Dim oMbrPart As New StructDetailObjects.MemberPart
        Set oMbrPart.object = oPortBounded.Connectable
        dBoundedFlgThickness = oMbrPart.flangeThickness
    ElseIf TypeOf oPortBounded.Connectable Is IJProfile Then
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = oPortBounded.Connectable
        dBoundedFlgThickness = oProfilePart.flangeThickness
    Else
    
    End If

    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    bTFL = False: bBFL = False: bTFR = False: bBFR = False

    CrossSection_Flanges oPortBounded.Connectable, bTFL, bBFL, bTFR, bBFR
    
    Dim dTopFlangeThickness As Double
    Dim dBottomFlangeThickness As Double
    dTopFlangeThickness = 0.00001 '0.01 mm
    dBottomFlangeThickness = 0.00001 '0.01 mm
     
    If (bTFL) Or (bTFR) Then dTopFlangeThickness = dBoundedFlgThickness
    If (bBFL) Or (bBFR) Then dBottomFlangeThickness = dBoundedFlgThickness
     
    pPRL.Add "FaceOffset", 0.00003
    pPRL.Add "TopFlangeThickness", dTopFlangeThickness
    pPRL.Add "BottomFlangeThickness", dBottomFlangeThickness
    pPRL.Add "TopOffset", 0.00003
    pPRL.Add "BottomOffset", 0.00003
    pPRL.Add "TopEdgeOffset", 0.00003
    pPRL.Add "TopEdgeInsideOffset", 0.00003
    pPRL.Add "TopEdgeOutsideOffset", 0.00003
    pPRL.Add "TopFaceOffset", 0.00003
    
    'define the cutting depth to completely cut the Member cross Section
    dCuttingDepth = EndCut_GetCutDepth(pPRL)
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



