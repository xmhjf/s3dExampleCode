' DI-CP-235020  MF: Bulkload errors on Oracle  
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMBPSFlangeParm"
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

Const m_ParameterRuleProgid As String = "SMBcktFlangeRules.SMBPSFlangeParm"
Const m_ParameterRuleName   As String = "SMBcktFlangeRules.SMBPSFlangeParm"
Const m_FamilyProgid        As String = ""

Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketFlangeRules\SMBPSFlangeParm.cls"

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

  pOH.SetOutput "NeutralAxisRatio"
  pOH.SetOutput "BendAngle"
  pOH.SetOutput "FlangeWidth"
  pOH.SetOutput "InnerRadius"
  pOH.SetOutput "RollLine1"
  pOH.SetOutput "RollLine2"
  pOH.SetOutput "KnuLine"
  
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
    
    Dim oFlangeContourSO As IJSmartOccurrence
    Dim oPlateUtils As IJBracketAttributes
    Dim oSOAttrs As IJDAttributes

    Dim dPlateThickness   As Double
    Dim dBendAngle        As Double
    Dim dFlangeWidth      As Double
    Dim dNAR              As Double
    Dim dInnerRadius      As Double
    Dim dOuterRadius      As Double
    Dim dNeutralRadius    As Double
    Dim dPI               As Double
    Dim dBendAngleRad     As Double
    Dim dSweepAngle       As Double
    Dim dArcLength        As Double
    Dim dRollLine1        As Double
    Dim dRollLine2        As Double
    Dim dKnuLine          As Double
    Dim dKnu2Edge         As Double
    Dim oAttrHelper       As IJDAttributes
    Dim oAttr             As IJDAttribute
    Dim oAttributeCol     As IMSAttributes.IJDAttributesCol
    
    dPI = 4 * Atn(1)
    
    Set oFlangeContourSO = pPRL.SmartOccurrence
    Set oAttrHelper = oFlangeContourSO
    
    ' Create the utility that will be used to get the bracket plate object from the flange
    ' smart occurrence.
    Dim oStrEntUtils As IJStructEntityUtils
    Set oStrEntUtils = New StructEntityUtils
    
    ' Get the bracket plate object.
    Dim oBracketPlate As IJPlate
    Set oBracketPlate = oStrEntUtils.GetFlangedPlate(pPRL.SmartOccurrence)
    
    ' always use the actual bracket thickness
    dPlateThickness = oBracketPlate.thickness
    
    'Get InnerRadius, NAR, bendangle and flange width from catalog...
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracketFlangeContrProp")
    Set oAttr = oAttributeCol.Item("NeutralAxisRatio")
    dNAR = oAttr.Value
    Set oAttr = Nothing
    
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracketFlangeContrProp")
    Set oAttr = oAttributeCol.Item("BendAngle")
    dBendAngle = oAttr.Value
    Set oAttr = Nothing
    
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracketFlangeContrProp")
    Set oAttr = oAttributeCol.Item("InnerRadius")
    dInnerRadius = oAttr.Value
    Set oAttr = Nothing
    
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracketFlangeContrProp")
    Set oAttr = oAttributeCol.Item("FlangeWidth")
    dFlangeWidth = oAttr.Value
    Set oAttr = Nothing
    
    'Compute bend lines...
    dOuterRadius = dInnerRadius + dPlateThickness
    dNeutralRadius = dInnerRadius + (dPlateThickness * dNAR)
    dSweepAngle = dPI - dBendAngle
    dArcLength = dSweepAngle * dNeutralRadius
    dKnu2Edge = dFlangeWidth - (dNeutralRadius * Tan(dSweepAngle / 2))
    dRollLine1 = dArcLength / 2
    dRollLine2 = dArcLength / 2
    dKnuLine = dKnu2Edge + dArcLength
    
    'set properties on the Flanged Bracket System - PlateFlange_AE
    '*************** Do Not modify - this call will bend the 3D shape in the model...*********************
    Set oPlateUtils = New PlateUtils
    oPlateUtils.SetPropertiesOnPlateFlangeAE oFlangeContourSO, dBendAngle, dInnerRadius, dPlateThickness
    '*****************************************************************************************************
    
    'Push outputs to symbol..
    pPRL.Add "NeutralAxisRatio", dNAR
    pPRL.Add "BendAngle", dBendAngle
    pPRL.Add "FlangeWidth", dFlangeWidth
    pPRL.Add "InnerRadius", dInnerRadius
    pPRL.Add "RollLine1", dRollLine1
    pPRL.Add "RollLine2", dRollLine2
    pPRL.Add "KnuLine", dKnuLine
        
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
' *******************************************************************************************
'         !!!!! End Private Code !!!!!
' *******************************************************************************************




