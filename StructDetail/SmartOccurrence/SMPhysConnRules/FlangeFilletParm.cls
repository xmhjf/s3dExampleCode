VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeFilletParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "PhysConnRules.FlangeFilletParm"
Const m_ParameterRuleName As String = CUSTOMERID + "PhysConnRules.FlangeFilletParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\FlangeFilletParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
  pOH.SetOutput "Nose"                      'User code "N"
  pOH.SetOutput "NoseMethod"
  pOH.SetOutput "NoseOrientationAngle"      'User code "n"
  pOH.SetOutput "MoldedFillet"
  pOH.SetOutput "AntiMoldedFillet"
  pOH.SetOutput "FilletMeasureMethod"
  pOH.SetOutput "Category"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  
  ' call sub to declare the computed TEE outputs for IJWeldSymbol
  AddWeldParmRuleOutputs FILLET_WELD1, pOH
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)

    On Error GoTo ErrorHandler
  
    Dim oParmCM As New PhysConnParmCM
    Dim dNoseOrientationAngle As Double
    Dim lNoseMethod As Long
    Dim dNose As Double
    Dim strRefPartName As String
    Dim dMoldedFillet As Double
    Dim dAntiMoldedFillet As Double
    Dim lFilletMeasureMethod As Long
    Dim strReferenceSide As String
    Dim lCategory As Long
    Dim bRefIsMolded As Boolean
    
    oParmCM.GetFilletWeldParameters oPRL, _
                                    "FlangeFillet", _
                                    False, _
                                    True, _
                                    dNoseOrientationAngle, _
                                    lNoseMethod, _
                                    dNose, _
                                    strRefPartName, _
                                    dMoldedFillet, _
                                    dAntiMoldedFillet, _
                                    lFilletMeasureMethod, _
                                    strReferenceSide, _
                                    lCategory, _
                                    bRefIsMolded

    ' Override fillets based on table
    Dim oPhysConn As StructDetailObjects.PhysicalConn
    Set oPhysConn = New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = oPRL.SmartOccurrence
    
    dMoldedFillet = oParmCM.GetFlangeFillet(oPhysConn.ConnectedObject1)
    dAntiMoldedFillet = dMoldedFillet
    
    oPRL.Add "NoseOrientationAngle", dNoseOrientationAngle
    oPRL.Add "NoseMethod", lNoseMethod
    oPRL.Add "Nose", dNose
    oPRL.Add "RefPartName", strRefPartName
    oPRL.Add "ReferenceSide", strReferenceSide
    oPRL.Add "MoldedFillet", dMoldedFillet
    oPRL.Add "AntiMoldedFillet", dAntiMoldedFillet
    oPRL.Add "FilletMeasureMethod", lFilletMeasureMethod
    oPRL.Add "Category", lCategory
    
    ' fill in the correct values for IJWeldSymbol
    ' this method will include a check for any of the input parameters from the bevel
    ' that have been overridden by the user
    SetCalculatedTeeWeldParams oPRL, FILLET_WELD1, bRefIsMolded, 0#, 0#, 0#, 0#, pcr_WG_None, dMoldedFillet, dAntiMoldedFillet, 0#, 0#, 0#
        
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
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
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************








