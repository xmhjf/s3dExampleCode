VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutNoseParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  File        : WebCutNoseParm.cls
'
'  Description : Parmameter rule for end cuts that need to define the cutting depth
'                 and need half of flange thickness value to offset WebCut point.
'
'  Author      : Alligators
'
'*********************************************************************************************

Const m_sClassName As String = "WebCutNoseParm"
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
    pOH.SetOutput "NoseOffsetDriven"
    pOH.SetOutput "FlangeThickness"
  
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
        
    Dim dCuttingDepth As Double
    dCuttingDepth = EndCut_GetCutDepth(oPRL)
    oPRL.Add "CutDepth", dCuttingDepth

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oPRL.SmartOccurrence
    
    Dim dNoseOffset As Double
    dNoseOffset = -1 'Initialize
    If TypeOf oSDO_WebCut.Bounding Is IJSmartPlate Then
        Dim eSmartPlateType As SmartPlateTypes
        Dim oSmartPlate As IJSmartPlate
        Set oSmartPlate = oSDO_WebCut.Bounding
        If oSmartPlate.SmartPlateType = SmartPlateTypes.spType_BEARING Then
            Dim eResultType As eResultType
            
            Dim oSDObjectHelper As New StructDetailObjects.Helper
            Dim strTableName As String
            Dim lCodeListValue As Long
            Dim strShortDesc As String
            Dim strLongDesc As String
            
            oSDObjectHelper.GetCustomAttributeParamValue oSDO_WebCut.Bounding, _
                                                        "IJUANoseOffset", _
                                                        "NoseOffsetDriving", _
                                                        dNoseOffset, _
                                                        eResultType, _
                                                         strTableName, _
                                                        lCodeListValue, _
                                                        strShortDesc, _
                                                        strLongDesc

        End If
    End If
    
    Dim dNoseOffsetD As Double 'Driven parameter
    dNoseOffsetD = dNoseOffset
    If LessThanOrEqualTo(dNoseOffsetD, 0) Then
        sMsg = sMsg & "NoseOffsetDrvg parameter is invalid: " & dNoseOffsetD
        GoTo ErrorHandler
    End If
    oPRL.Add "NoseOffsetDriven", dNoseOffsetD
    
    Dim dFlangeThickness As Double
    
    dFlangeThickness = 0.01 '10 mm 'Initial/default value
    If TypeOf oSDO_WebCut.Bounded Is ISPSMemberPartPrismatic Then
        Dim oSDO_BoundedMbr As StructDetailObjects.MemberPart
        Set oSDO_BoundedMbr = New StructDetailObjects.MemberPart
        Set oSDO_BoundedMbr.object = oSDO_WebCut.Bounded
        dFlangeThickness = oSDO_BoundedMbr.flangeThickness
    End If
    
    oPRL.Add "FlangeThickness", dFlangeThickness

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




