VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebSnipNoNoseParm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "EndCutRules.WebSnipNoNoseParm"
Const m_ParameterRuleName As String = CUSTOMERID + "EndCutRules.WebSnipNoNoseParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\WebSnipNoNoseParm.cls"

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
  pOH.SetOutput "SetBack"
  pOH.SetOutput "Slope"
  
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

     Dim oPortBounded As IJPort
     Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    
    'define the cutting depth as four(4) times the Web Thickness
    Dim dCuttingDepth As Double
    dCuttingDepth = GetBoundedCutDepth(oPortBounded.Connectable, False) * 4#
    pPRL.Add "CutDepth", dCuttingDepth
    
    'get answers from selector
    Dim sTheta As String
    'TR-169311 Trying to retrieve question's answer without
    'hardcoding the path
    GetSelectorAnswer pPRL, "Theta", sTheta
    
    'Get WebCut smart occurence to determine if this is a FreeEnd WebCut
    Dim bFreeEndCut As Boolean
    Dim oWebCut As New StructDetailObjects.WebCut
    Set oWebCut.object = pPRL.SmartOccurrence
    bFreeEndCut = oWebCut.IsFreeEndCut
    
    Dim dSetBack As Double
    If bFreeEndCut Then
        Dim bIsUnBounded As Boolean
        bIsUnBounded = oWebCut.IsUnBounded
        If bIsUnBounded Then
            ' FreeEndCut is bounded by edge of Stiffened Part(Plate)
            ' default is 35mm
            dSetBack = 0.035
        
        Else
            ' FreeEndCut is bounded by Ref. Plane, Assoc. Point, or
            ' Landing Curve definition (not by Stiffened Plate Edge)
            ' default is 0.01mm
            dSetBack = 0.00001
        End If

    Else
        ' default is 35mm
        dSetBack = 0.035
    End If
    
    'set the slope value, based on the type of profile and the length of the profile
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oWebCut.Bounded
    Dim sSectionType As String
    Dim dHeight As Double
    Dim dLength As Double
    Dim dSlope As Double
    
    sSectionType = oProfilePart.SectionType
    
    'get the value of theta
    Dim dTheta As Double
    dTheta = Val(sTheta)
    dTheta = DegToRad(dTheta)
    
    If dTheta >= DegToRad(65) And dTheta < DegToRad(115) Then
        'set the slope based on theta
        dSlope = dTheta
    Else
        'the value if out of range, so set a default
        dSlope = 0.314159 '18 degrees
    End If
        
    pPRL.Add "Slope", dSlope
    pPRL.Add "SetBack", dSetBack

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
'*****************************************************************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    Dim dSlope As Double
    Dim dOffset As Double
    
    dSlope = GetCustomAttribute(pSmartOccurrence, "IJUASmartEndCuts", "Slope")
    dOffset = GetCustomAttribute(pSmartOccurrence, "IJUAWebCuts", "SetBack")
    
    Dim sDrawingtype As String
    If GreaterThan(dOffset, 0.004) Then
        ' offset
        If IsValueBetween(dSlope, 0, degreeToRadian(89.99)) Or IsValueBetween(dSlope, degreeToRadian(90.01), degreeToRadian(180)) Then
            sDrawingtype = Sniped_Offset_WebCuts
        Else
            sDrawingtype = Straight_Offset_WebCuts
        End If
    Else
        If IsValueBetween(dSlope, 0, degreeToRadian(89.99)) Or IsValueBetween(dSlope, degreeToRadian(90.01), degreeToRadian(180)) Then
            sDrawingtype = Sniped_No_Offset_WebCuts
        Else
            sDrawingtype = Straight_No_Offset_WebCuts
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







