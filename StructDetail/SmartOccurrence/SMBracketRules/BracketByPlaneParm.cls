VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketByPlaneParm"
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

Const m_ParameterRuleProgid As String = m_sProjectName + ".BracketByPlaneParm"
Const m_ParameterRuleName As String = m_sProjectName + ".BracketByPlaneParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "BracketRules\BracketByPlaneParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_SUPPORT1
    pIH.SetInput INPUT_SUPPORT2
    pIH.SetInput INPUT_PLANE
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "Height1"
  pOH.SetOutput "Width1"
  pOH.SetOutput "Nose1"
  pOH.SetOutput "Nose2"

  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
  
    Dim oBracketContour As IJSmartOccurrence
    Dim oAttrHelper As IJDAttributes
    Dim oAttr As IJDAttribute
    Dim oAttributeCol As IMSAttributes.IJDAttributesCol
    
    Set oBracketContour = pPRL.SmartOccurrence
  
    Set oAttrHelper = oBracketContour

    'get the number of supports selected
    Dim nSupports As Long
    Dim oSupports As IJElements
    
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
   
    'Get the number of supports from BracketContourSO.
    oBracketUtils.GetSupportsFromBracketContourSO oBracketContour, oSupports, nSupports

    Dim flangewidth As Double
    Dim webheight As Double

    
    
    ' The attributes Height1,Width1,Nose1 and Radius1 are availabe from
    ' the interface IJUABracketProps.Get them from catalog and set on the SmartItem
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracketProps")
    Dim s1DimName As String
    Dim s2DimName As String
    Dim s1DimKey As String
    Dim s2DimKey As String
    s1DimName = "Width1"
    s2DimName = "Height1"
    If nSupports = 2 Then
        Dim bracketSymbol As IJDSymbol
        Set bracketSymbol = oBracketContour
    
        oBracketUtils.FindDimensionConstraint bracketSymbol.IJDSymbolDefinition(2), "support1", s1DimKey
        oBracketUtils.FindDimensionConstraint bracketSymbol.IJDSymbolDefinition(2), "support2", s2DimKey
        Dim anInput As IJDInput
        For Each anInput In bracketSymbol.IJDSymbolDefinition(2).IJDInputs
            If anInput.Key = s1DimKey Then
                s1DimName = anInput.Name
            ElseIf anInput.Key = s2DimKey Then
                s2DimName = anInput.Name
            End If
        Next
        Dim oProfileObject As StructDetailObjects.ProfileSystem
        Set oProfileObject = New StructDetailObjects.ProfileSystem
        Dim oDCrossSection As IJDProfileSection
        With oProfileObject
        ' use the affected leaf system as the cross section can change
        Set .object = oSupports(1)
        Set oDCrossSection = .object
        'if the profile has a flange, it can be used for a "1S"
        If (oDCrossSection.SectionType And Tee) > 0 Then
            flangewidth = .FlangeLength * 0.4 'want the bracket to be 40% of the flange
        Else
            flangewidth = .FlangeLength * 0.8 'want the bracket to be 80% of the flange
    End If
    
        webheight = .WebLength - .FlangeThickness
    End With

    End If
    
    Dim paramVal As Double
    Dim scaleFactor As Double
    scaleFactor = 1#
    Set oAttr = oAttributeCol.Item(s2DimName)
    paramVal = oAttr.Value
    If paramVal > flangewidth And (nSupports = 2) Then
        scaleFactor = flangewidth / paramVal
        paramVal = flangewidth
    End If

    pPRL.Add s2DimName, paramVal
    
    Set oAttr = Nothing
    
    Set oAttr = oAttributeCol.Item(s1DimName)
    paramVal = oAttr.Value * scaleFactor
    If (paramVal > webheight) And (nSupports = 2) Then
            paramVal = webheight
    End If
    pPRL.Add s1DimName, paramVal
    
   
    
    
    Set oAttr = Nothing
    Set oAttr = oAttributeCol.Item("Nose1")
    pPRL.Add "Nose1", oAttr.Value
    
    Set oAttr = Nothing
    Set oAttr = oAttributeCol.Item("Radius1")
    pPRL.Add "Radius1", oAttr.Value
        
    ' The attributes Height2,Width2,Nose2 and Radius2 are availabe from
    ' the interface IJUABracketProps. Get them from catalog and set on the SmartItem
    Set oAttributeCol = Nothing
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracket2Props")
  
    Set oAttr = Nothing
    Set oAttr = oAttributeCol.Item("Height2")
    pPRL.Add "Height2", oAttr.Value
    
    Set oAttr = Nothing
    Set oAttr = oAttributeCol.Item("Width2")
    pPRL.Add "Width2", oAttr.Value
    
    Set oAttr = Nothing
    Set oAttr = oAttributeCol.Item("Nose2")
    pPRL.Add "Nose2", oAttr.Value
    
    Set oAttr = Nothing
    Set oAttr = oAttributeCol.Item("Radius2")
    pPRL.Add "Radius2", oAttr.Value
      
    Set oAttr = Nothing
    Set oAttributeCol = Nothing
    
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


