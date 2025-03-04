VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqrAlgFlgRelBotCutSnpParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : SqrAlgFlgRelBotCutSnpParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    22/Sep/2011 - svsmylav TR-202526: corrected logic to get BottomShapeOutside answer instead of TopShapeOutside.
'    07/Oct/2011 - svsmylav TR-202529: adjust Radius and Edge-Offset parameter values to avoid zero length tangent line.
'    19/Oct/2011 - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'
'    11/Apr/2012 - svsmylav
'       DM-213229: Updated logic to handle stiffener as bounding object.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************
Const m_sClassName As String = "SqrAlgFlgRelBotCutSnpParm"
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
    pOH.SetOutput "FlangeThickness"
    pOH.SetOutput "Slope"
    pOH.SetOutput "SlopeOutside"
    pOH.SetOutput "FlangeClearance"
    pOH.SetOutput "BottomEdgeOffset"
    pOH.SetOutput "Nose"
    pOH.SetOutput "BottomEdgeInsideOffset"
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
Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
    
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    Dim dRadius As Double
    Dim dFlangeThickness As Double
    Dim dEdgeInsideOffset As Double
    Dim dEdgeOffset As Double
    Dim dCutDepth As Double
    Dim dEdgeLength As Double
    Dim dMaxNose As Double
    
    Dim oParmCM As New EndCutParmCM
    oParmCM.GetSquareFlangeReliefParameters oPRL, _
                                            True, _
                                            dRadius, _
                                            dFlangeThickness, _
                                            dEdgeInsideOffset, _
                                            dEdgeOffset, _
                                            dCutDepth, _
                                            dEdgeLength, _
                                            dMaxNose
                                                         
    Dim dSlope As Double
    Dim dSlopeOutside As Double
    Dim dFlangeClearance As Double
    Dim dNose As Double
                                                         
    oParmCM.GetOutsideCutSnipeParameters oPRL, True, dSlope, dSlopeOutside, dFlangeClearance, dNose
    
    ' ---------------------
    ' Adjust the nose value
    ' ---------------------
    ' The method above assumes the nose starts at the top of the bounding flange
    ' We want the nose to start at the end of the radius, if it extends above the bounding flange, as long as there is enough room
    If dRadius - dEdgeInsideOffset > dEdgeLength Then
        dNose = dNose + dRadius - dEdgeInsideOffset - dEdgeLength
    End If
    
    If dNose >= dMaxNose Then
        dNose = dRadius - dEdgeInsideOffset - dEdgeLength + 0.00001
    End If
        
    oPRL.Add "Radius", dRadius
    oPRL.Add "FlangeThickness", dFlangeThickness
    oPRL.Add "BottomEdgeOffset", dEdgeOffset
    oPRL.Add "BottomEdgeInsideOffset", dEdgeInsideOffset
    oPRL.Add "CutDepth", dCutDepth
    oPRL.Add "Slope", dSlope
    oPRL.Add "SlopeOutside", dSlopeOutside
    oPRL.Add "FlangeClearance", dFlangeClearance
    oPRL.Add "Nose", dNose

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

' ********************************************************************************************
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

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
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
    Dim dSlope As Double
    Dim dEdgeOffset As Double
    
    dSlope = GetCustomAttribute(pSmartOccurrence, "IJUAMbrFreeEndCuts", "Slope")
    dEdgeOffset = GetCustomAttribute(pSmartOccurrence, "IJUAMbrEndCut", "BottomEdgeOffset")
    
    If IsValueBetween(dSlope, 0, degreeToRadian(89.99)) Or IsValueBetween(dSlope, degreeToRadian(90.01), degreeToRadian(180)) Then
        If GreaterThan(dEdgeOffset, 0.004) Then
            IJEndCutDrawingType_UpdateDrawingType = Sniped_Offset_WebCuts
        Else
            IJEndCutDrawingType_UpdateDrawingType = Sniped_No_Offset_WebCuts
        End If
    Else
        If GreaterThan(dEdgeOffset, 0.004) Then
            IJEndCutDrawingType_UpdateDrawingType = Straight_Offset_WebCuts
        Else
            IJEndCutDrawingType_UpdateDrawingType = Straight_No_Offset_WebCuts
        End If
    End If
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.BottomOrRight, IJEndCutDrawingType_UpdateDrawingType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
