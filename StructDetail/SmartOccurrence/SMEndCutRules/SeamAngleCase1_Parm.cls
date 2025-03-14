VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamAngleCase1_Parm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "EndCutRules.SeamAngleCase1_Parm"
Const m_ParameterRuleName As String = CUSTOMERID + "EndCutRules.SeamAngleCase1_Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\SeamAngleCase1_Parm.cls"

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType


Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "CutDepth"
  pOH.SetOutput "WebSeamAngle"
  pOH.SetOutput "WebSeamPointOffset"
  pOH.SetOutput "FlangeSeamAngle"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

    Dim oPortBounded As IJPort
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
     
    'Get Bounded part CutDepth
    Dim dCuttingDepth As Double
    dCuttingDepth = GetBoundedCutDepth(oPortBounded, True)
    
'    ' get the part number;
'    Dim sPartNumber As String
'    sPartNumber = pPRL.SelectorAnswer(CUSTOMERID + "EndCutRules.EndToEndWebCutSel", _
'                                      gsWeldPart)
    
    ' want the WebCut to also Cut the Flange
    ' set Cutting Depth based on Flange Length
    dCuttingDepth = dCuttingDepth * 3#
    pPRL.Add "CutDepth", dCuttingDepth
    
    'Second Symbol contains angles measured opposite of First Symbol
    'Therefore Input values are same for both First, Second Symbols
    Dim dPI As Double
    dPI = Atn(1#) * 4

    ' If lapped connection, copy angle between bounding and bounded stiffeners
    Dim dLapDist As Double
    If IsLapped(oPortBounded, dLapDist) Then
        Dim angle As Double
        angle = GetBoundingAngle(pPRL.SmartOccurrence)
        pPRL.Add "WebSeamAngle", angle
        pPRL.Add "FlangeSeamAngle", dPI / 2#
        pPRL.Add "WebSeamPointOffset", dLapDist
    ' If split connection...
    Else
        'Second Symbol contains angles measured opposite of First Symbol
        'Therefore Input values are same for both First, Second Symbols
        pPRL.Add "WebSeamAngle", dPI / 3#
        pPRL.Add "FlangeSeamAngle", dPI / 2#
        pPRL.Add "WebSeamPointOffset", 0.00001
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
'**********************************************************
' Method Name : IJEndCutDrawingType_UpdateDrawingType
' Inputs      : feature occurence object
' Output      : string, returns the end cut type.
' Description : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*********************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    Dim dOffset As Double
    Dim dAngle As Double
    
    dOffset = GetCustomAttribute(pSmartOccurrence, "IJUAAngledEndCuts", "WebSeamPointOffset")
    dAngle = GetCustomAttribute(pSmartOccurrence, "IJUAAngledEndCuts", "WebSeamAngle")
    
    Dim sDrawingtype As String
    
    If GreaterThan(dOffset, 0.004) Then
        ' offset
        If IsValueBetween(dAngle, 0, degreeToRadian(89.99)) Or IsValueBetween(dAngle, degreeToRadian(90.01), degreeToRadian(180)) Then
            sDrawingtype = Sniped_Offset_WebCuts
        Else
            sDrawingtype = Straight_Offset_WebCuts
        End If
    Else
        If IsValueBetween(dAngle, 0, degreeToRadian(89.99)) Or IsValueBetween(dAngle, degreeToRadian(90.01), degreeToRadian(180)) Then
            sDrawingtype = Sniped_No_Offset_WebCuts
        Else
            sDrawingtype = Straight_No_Offset_WebCuts
        End If
    End If
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, sDrawingtype
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
    
End Function

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************






