VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BktByPlaneFlParm"
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

Const m_ParameterRuleProgid As String = "SMBcktFlangeRules.BktByPlaneFlParm"
Const m_ParameterRuleName   As String = "SMBcktFlangeRules.BktByPlaneFlParm"
Const m_FamilyProgid        As String = ""

Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketFlangeRules\BracketByPlaneFlParm.cls"

Private METHOD As String

Implements IJDUserSymbolServices
' *******************************************************************************************
Public Sub ParameterRuleInputs( _
    pIH As IJDInputsHelper _
)

  On Error GoTo ErrorHandler
  METHOD = "ParameterRuleInputs"

  'what are the inputs?
  
  Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
' *******************************************************************************************
Public Sub ParameterRuleOutputs( _
    pOH As IJDOutputsHelper _
)
  On Error GoTo ErrorHandler
  METHOD = "ParameterRuleOutputs"

  pOH.SetOutput "Tan2Knu"
  pOH.SetOutput "Knu2Edge"
  
  Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
' *******************************************************************************************
Public Sub ParameterRuleLogic( _
    pPRL As IJDParameterLogic _
)
    On Error GoTo ErrorHandler
    METHOD = "ParameterRuleLogic"

    'get the parameter values that are in the catalog, or over-ridden by user
    Dim dPlateThickness   As Double
    Dim dBendAngleDegrees As Double
    Dim dFlangeWidth      As Double
    Dim dMoldedSurface    As Double
    Dim dInnerRadius      As Double
        
    dPlateThickness = pPRL.SmartItemAttribute("PlateThickness")
    dBendAngleDegrees = pPRL.SmartItemAttribute("BendAngle")
    dFlangeWidth = pPRL.SmartItemAttribute("FlangeWidth")
    dMoldedSurface = pPRL.SmartItemAttribute("MoldedSurfaceSide")
    dInnerRadius = pPRL.SmartItemAttribute("InnerRadius")
      
    Dim oFlangeCountourSO As IJSmartOccurrence
    Dim oPlateUtils As IJBracketAttributes
    Set oPlateUtils = New PlateUtils
    
    Set oFlangeCountourSO = pPRL.SmartOccurrence
    
    'set properties on the Flanged Bracket System - PlateFlange_AE
    oPlateUtils.SetPropertiesOnPlateFlangeAE oFlangeCountourSO, dBendAngleDegrees, dInnerRadius, dPlateThickness
    
    'Convert the bendangle to radians, and actually compute the complement
    'which we need in the formulas.
'    Const Pi  As Double = 3.141592654
'    Dim angle As Double
'    angle = Pi - dBendAngleDegrees * Pi / 180
    
    'calculate the outputs
    Dim dTan2Knu  As Double
    Dim dKnu2Edge As Double
          
    dTan2Knu = (dBendAngleDegrees / 2) * (dInnerRadius + (dPlateThickness * dMoldedSurface))
    dKnu2Edge = dFlangeWidth - ((dInnerRadius + dPlateThickness) * Sin(dBendAngleDegrees)) + dTan2Knu
  
    pPRL.Add "Tan2Knu", dTan2Knu
    pPRL.Add "Knu2Edge", dKnu2Edge
        
  Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **





' *******************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' *******************************************************************************************
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
' *******************************************************************************************
'         !!!!! End Private Code !!!!!
' *******************************************************************************************


