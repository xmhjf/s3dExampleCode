VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeCutParm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "EndCutRules.FlangeCutParm"
Const m_ParameterRuleName As String = CUSTOMERID + "EndCutRules.FlangeCutParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\FlangeCutParm.cls"

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
  pOH.SetOutput "BottomFlange"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

     Dim oPortBounded As IJPort
     Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
     Dim oBounded As Object
     Dim phelper As New StructDetailObjects.Helper
     
     Dim dBoundedFlangeThickness As Double
     
    'Get Bounded part
    If phelper.ObjectType(oPortBounded.Connectable) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oPortBounded.Connectable
        Set oBounded = oBoundedProfile
        dBoundedFlangeThickness = oBoundedProfile.FlangeThickness
    ElseIf phelper.ObjectType(oPortBounded.Connectable) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oPortBounded.Connectable
        Set oBounded = oBoundedBeam
        dBoundedFlangeThickness = oBoundedBeam.FlangeThickness
    End If
    
    'define the cutting depth as one-half the Web Length
    'this assumes
    ' the Top and Bottom Flanges are symmetrical along the Web Length
    'and
    ' The Flange's Thickness is never over Half the Web Length
    
    'Get the Smart Item Name
    Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
    Set oFlangeCutWrapper.object = pPRL.SmartOccurrence
    
    Dim oSmartOccurrence As IJSmartOccurrence
    Set oSmartOccurrence = oFlangeCutWrapper.object
    
    Dim oSmartItem As IJSmartItem
    Dim sItemName As String
    
    Set oSmartItem = oSmartOccurrence.SmartItemObject
    sItemName = oSmartItem.Name
    
    Dim dCuttingDepth As Double
       
    If sItemName = "FlangeCut_W1" Or sItemName = "FlangeCut_W4" Then
      dCuttingDepth = dBoundedFlangeThickness
    Else
     dCuttingDepth = oBounded.WebLength * 0.6
    End If

    pPRL.Add "CutDepth", dCuttingDepth

    Dim sBottomFlange As String
    GetParentAnswer pPRL, gsBottomFlange, sBottomFlange

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
'**********************************************************
' Method Name   : IJEndCutDrawingType_UpdateDrawingType
' Inputs        : feature occurence object
' Output        : string, returns the end cut type.
' Description   : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*********************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    Dim sDrawingtype As String
    
    Dim dSlope As Double
    Dim dProfilepointOffset As Double
    
    dSlope = GetCustomAttribute(pSmartOccurrence, "IJUASmartEndCuts", "Slope")
    dProfilepointOffset = GetCustomAttribute(pSmartOccurrence, "IJUAFlangeCuts", "ProfilePointOffset")
    
    If GreaterThan(dProfilepointOffset, 0.004) Then
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



