VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToLongBoxFlCutParm"
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
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "EndCutRules.EndToLongBoxFlCutParm"
Const m_ParameterRuleName As String = CUSTOMERID + "EndCutRules.EndToLongBoxFlCutParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\EndToLongBoxFlCutParm.cls"

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType


'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "CutDepth"
  pOH.SetOutput "BottomFlange"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

    Dim dCutDepth As Double
    Dim sBottomFlange As String
    Dim oPortBounded As IJPort
    Dim oBoundedProfile As StructDetailObjects.ProfilePart
    
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
     
    'Get Bounded part
    Set oBoundedProfile = New StructDetailObjects.ProfilePart
    Set oBoundedProfile.object = oPortBounded.Connectable
    
    dCutDepth = oBoundedProfile.WebLength * 0.6
    pPRL.Add "CutDepth", dCutDepth

    'TR-169311 Trying to retrieve question's answer without
    'hardcoding the path
    GetSelectorAnswer pPRL, gsBottomFlange, sBottomFlange
    
    If sBottomFlange = "No" Then
        pPRL.Add "BottomFlange", 0
    Else
        pPRL.Add "BottomFlange", 1
    End If

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
'*****************************************************************************************************
' Method Name : IJEndCutDrawingType_UpdateDrawingType
' Inputs      : feature occurence object
' Output      : string, returns the end cut type.
' Description : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*********************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    
    On Error GoTo ErrorHandler
    
    Dim dSlope As Double
    Dim dOffset As Double
    Dim sDrawingtype As String
    
    dSlope = GetCustomAttribute(pSmartOccurrence, "IJUASmartEndCuts", "Slope")
    dOffset = GetCustomAttribute(pSmartOccurrence, "IJUAFlangeCuts", "ProfilePointOffset")

    If GreaterThan(dOffset, 0.004) Then
        ' offset
        If IsValueBetween(dSlope, 0, degreeToRadian(89.99)) Or IsValueBetween(dSlope, degreeToRadian(90.01), degreeToRadian(180)) Then
            sDrawingtype = Sniped_Offset_FlangeCuts
        Else
            sDrawingtype = Straight_Offset_FlangeCuts
        End If
    Else
        If IsValueBetween(dSlope, 0, degreeToRadian(89.99)) Or IsValueBetween(dSlope, degreeToRadian(90.01), degreeToRadian(180)) Then
            sDrawingtype = Sniped_No_Offset_FlangeCuts
        Else
            sDrawingtype = Straight_No_Offset_FlangeCuts
        End If
    End If
        
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, sDrawingtype
    IJEndCutDrawingType_UpdateDrawingType = sDrawingtype
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
    
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




