VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bracket4SProfileParm"
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

Const m_ParameterRuleProgid As String = "BracketRules.Bracket4SProfileParm"
Const m_ParameterRuleName As String = "BracketRules.Bracket4SProfileParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\Bracket4SProfileParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_SUPPORT1
    pIH.SetInput INPUT_SUPPORT2
    pIH.SetInput INPUT_SUPPORT3
    pIH.SetInput INPUT_SUPPORT4
    pIH.SetInput INPUT_PLANE
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "Scallop1"
  pOH.SetOutput "Scallop2"
  pOH.SetOutput "Scallop3"
  pOH.SetOutput "Offset1"
  pOH.SetOutput "Offset2"
  pOH.SetOutput "Radius1"
  pOH.SetOutput "Radius2"
  pOH.SetOutput "Offset1a"
  pOH.SetOutput "Offset2a"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler

    Dim nSupports As Long
    Dim oSupportsColl As IJElements
    Dim oBracketContour As IJSmartOccurrence
    Dim oAttrHelper As IJDAttributes
    Dim oAttr As IJDAttribute
    Dim oAttributeCol As IMSAttributes.IJDAttributesCol
    Dim dProfileThk As Double
    Dim oDCrossSection As IJDProfileSection
    Dim dScallop1, dScallop2, dScallop3 As Double
    Dim dRadius1, dRadius2 As Double
    Dim dOffset1, dOffset2 As Double
    Dim dOffset1a, dOffset2a As Double
    
    ' Get Bracket-by-Plane Smart Occurrence:
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Set oBracketByPlaneSO = pPRL.SmartOccurrence
    
    ' Create helper for final construction of bracket:
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    
    'Get the supports from BracketContourSO.
    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupportsColl, nSupports
    
    'Get Supports
    Dim oSupport3 As Object
    Dim oSupport4 As Object

    If Not oSupportsColl Is Nothing Then
        If oSupportsColl.Count > 3 Then
            Set oSupport3 = oSupportsColl.Item(3)
            Set oSupport4 = oSupportsColl.Item(4)
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
      'Check support  4 and 5 is plate or profile
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    Dim oProfileSup As StructDetailObjects.ProfileSystem
    
    Select Case oBracketByPlaneSO.Item
        Case "4SCLTOffSetFromFlangeScallop"
                'Both flanges are pointing towards bracket system (IN)
                    dScallop1 = 0.05
                    dScallop2 = 0.035
                    dScallop3 = 0.035
                    dRadius1 = 0.1
                    dRadius2 = 0.1
                    dOffset1 = 0.015
                    dOffset1a = dOffset1 + 0.02
                    dOffset2 = 0.015
                    dOffset2a = dOffset2 + 0.02
                    
        Case "4SCLTOffSetFromSup3FlangeInSc"
                Set oProfileSup = New StructDetailObjects.ProfileSystem
                Set oProfileSup.object = oSupportsColl.Item(4)
                
                'Support 3 is a flat bar (no flange) or a plate system...
                If (oProfileSup.SectionType = "FB") Or (pHelper.ObjectType(oSupport3) = SDOBJECT_PLATESYSTEM) Then
                       dScallop1 = 0.05
                       dScallop2 = 0.035
                       dScallop3 = 0.035
                       dRadius1 = 0.075
                       dRadius2 = 0.075
                       dOffset1 = 0.015
                       dOffset1a = dOffset1 + 0.06
                       dOffset2 = 0.015
                       dOffset2a = dOffset2 + 0.02
                        
                Else 'Then the support has a flange with direction same as bracket
                       dScallop1 = 0.05
                       dScallop2 = 0.035
                       dScallop3 = 0.035
                       dRadius1 = 0.1
                       dRadius2 = 0.1
                       dOffset1 = 0.02
                       dOffset1a = dOffset1 + 0.02
                       dOffset2 = 0.015
                       dOffset2a = dOffset2 + 0.02

                End If

        Case "4SCLTOffSetFromSup4FlangeInSc"
                
                Set oProfileSup = New StructDetailObjects.ProfileSystem
                Set oProfileSup.object = oSupportsColl.Item(4)
                
                'Support 4 is a flat bar (no flange) or a plate system...
                If (oProfileSup.SectionType = "FB") Or (pHelper.ObjectType(oSupport4) = SDOBJECT_PLATESYSTEM) Then
                       dScallop1 = 0.05
                       dScallop2 = 0.035
                       dScallop3 = 0.035
                       dRadius1 = 0.075
                       dRadius2 = 0.075
                       dOffset1 = 0.015
                       dOffset1a = dOffset2 + 0.02
                       dOffset2 = 0.015
                       dOffset2a = dOffset1 + 0.06
                        
                Else ''Then the support has a flange with direction same as bracket
                       dScallop1 = 0.05
                       dScallop2 = 0.035
                       dScallop3 = 0.035
                       dRadius1 = 0.075
                       dRadius2 = 0.075
                       dOffset1 = 0.015
                       dOffset1a = dOffset1 + 0.02
                       dOffset2 = 0.02
                       dOffset2a = dOffset2 + 0.02

                End If

    End Select

    '******************************************************
    ' Set values...
    pPRL.Add "Scallop1", dScallop1
    pPRL.Add "Offset1", dOffset1
    pPRL.Add "Radius1", dRadius1
    pPRL.Add "Radius2", dRadius2
    pPRL.Add "Scallop2", dScallop2
    pPRL.Add "Scallop3", dScallop3
    pPRL.Add "Offset2", dOffset2
    pPRL.Add "Offset1a", dOffset1a
    pPRL.Add "Offset2a", dOffset2a

    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
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
  pIH.Definition = pPR
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



