VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014-16, Intergraph Corporation.  All rights reserved.
'
'  File        : WebCutParm.cls
'
'  Description : Parmameter rule for end cuts that need only to define the cutting depth
'
'  Author      : Alligators

' 03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
' 11/Jun/15     svsmylav   CR-271082 Create new Std. AC for 'ToEdgecase'.
' 6/Jan/16    pyk          DI-CP-273985 OM0001 offset incorrect when bounding top/bottom faces bounded part
'*********************************************************************************************

Const m_sClassName As String = "WebCutParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

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
    
    pOH.SetOutput "CutDepth"
    pOH.SetOutput "WebClearance"
  
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
    ' --------------
    ' Get the inputs
    ' --------------
    Dim oPortBounding As IJPort
    Dim oPortBounded As Object
    Set oPortBounding = oPRL.InputObject(INPUT_BOUNDING)
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)
    
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oPortBounding, oPortBounded)
    
    Dim dCuttingDepth As Double
    If bPenetratesWeb Then
        'Web-penetrated case
        dCuttingDepth = EndCut_GetCutDepth(oPRL)
    Else
        'Flange-penetrated case: WebCut cutting depth should be limited to
        ' the bounded part web thickness (FlangeCuts would remove remaining),
        ' however for rectangular member it would be complete Width
        If TypeOf oPortBounded Is IJPort Then
            If TypeOf oPortBounded.Connectable Is ISPSMemberPartCommon Then
                If IsRectangularMember(oPortBounded.Connectable) Then
                    dCuttingDepth = EndCut_GetCutDepth(oPRL)
                Else
                    Dim oBoundedMemberPart As StructDetailObjects.MemberPart
                
                    'Define the cutting depth as Web thickness + Small Tolearcane(0.1mm) so that
                    'it comepletely cuts the Web but not complete cross section and touches the Flange Notch
                    Set oBoundedMemberPart = New MemberPart
                    Set oBoundedMemberPart.object = oPortBounded.Connectable
                    dCuttingDepth = oBoundedMemberPart.webThickness + 0.0001
                    Set oBoundedMemberPart = Nothing
                End If
            End If
        End If
    End If
     ' ---------------------
    ' Set the cutting depth
    ' ---------------------
    oPRL.Add "CutDepth", dCuttingDepth
    
    Dim oBoundingObject As Object
    Set oBoundingObject = oPortBounding.Connectable
   
    Dim OSDO_MemberPart As New StructDetailObjects.MemberPart
    If TypeOf oBoundingObject Is ISPSMemberPartCommon Then
        Set OSDO_MemberPart.object = oBoundingObject
    End If
        
    Dim eCase As eMemberBoundingCase
    eCase = GetMemberBoundingCase(oPRL.SmartOccurrence, , , False)
     ' -------------------------------------------------------
    ' Get the port that maps to top or bottom, as appropriate
    ' -------------------------------------------------------
  '  Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Set oEdgeMapColl = GetEdgeMap(oPRL.SmartOccurrence, oPortBounding, oPortBounded, , bPenetratesWeb)
    
    Dim mapEdgeID As JXSEC_CODE
    mapEdgeID = JXSEC_UNKNOWN
    
    Dim realEdgeID As JXSEC_CODE
    realEdgeID = JXSEC_UNKNOWN
    
    If eCase = OnMemberBottom Then
        mapEdgeID = JXSEC_BOTTOM
    ElseIf eCase = OnMemberTop Then
        mapEdgeID = JXSEC_TOP
    End If
    
    Dim dClearance As Double
    
    If (ItemExists(mapEdgeID, oEdgeMapColl, realEdgeID)) Then
    
            Select Case realEdgeID
                ' ---------------------------------------------------------------------------
                ' If the port mapped is Top or Bottom, the clearance is half the flange width
                ' ---------------------------------------------------------------------------
                ' Presuming that for now, separate top and bottom flange widths are not suppored by the API
                Case JXSEC_TOP, JXSEC_BOTTOM
                    
                    If Not OSDO_MemberPart Is Nothing Then
                        dClearance = OSDO_MemberPart.FlangeLength / 2#
                    Else
                        dClearance = 0.00003
                    End If
                
                ' ----------------------------------------------------------------------------------
                ' If the port mapped is WebRight or WebLeft, the clearance is half the section depth
                ' ----------------------------------------------------------------------------------
                Case JXSEC_WEB_RIGHT, JXSEC_WEB_LEFT
                
                    If Not OSDO_MemberPart Is Nothing Then
                        dClearance = OSDO_MemberPart.Height / 2#
                    Else
                        dClearance = 0.00003
                    End If
                
                ' ------------------------------------------
                ' Other faces are not supported - use 0.03mm
                ' ------------------------------------------
                Case Else
                
                    dClearance = 0.00003
            
            End Select
    Else
        dClearance = 0.00003
    End If
    
    oPRL.Add "WebClearance", dClearance
   
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
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionProgid
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
  Dim oPRL As IJDParameterLogic
  Set oPRL = New ParameterLogic
  oPRL.Representation = pRep
  ParameterRuleLogic oPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




