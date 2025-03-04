VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BPRibConnParm"
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

Const m_ParameterRuleProgid As String = m_sProjectName & ".BPRibConnParm"
Const m_ParameterRuleName As String = m_sProjectName & ".BPRibConnParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SMBracketRules\BPRibConnParm.cls"

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
  pOH.SetOutput "Scallop"
  
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
    Dim paHeight, paWidth, paNose1, paNose2, paScallop As Double
    Dim dProfileWebLen, dProfileFlangeLen As Double
    Dim bSup1IsProfile As Boolean

    Set oBracketContour = pPRL.SmartOccurrence

    Set oAttrHelper = oBracketContour
    Dim nSupports As Long
    Dim oSupports As IJElements

    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils

    'Get the number of supports from BracketContourSO.
    oBracketUtils.GetSupportsFromBracketContourSO oBracketContour, oSupports, nSupports

    ' Find out which support is the bracket and get the section type of the bracket...
    Dim oSupport1 As Object
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    Set oSupport1 = oSupports(1)
    Dim oProfileObject As StructDetailObjects.ProfileSystem
    Set oProfileObject = New StructDetailObjects.ProfileSystem
    Dim oPlateObject As StructDetailObjects.PlatePart
    Set oPlateObject = New StructDetailObjects.PlatePart
    Dim oDCrossSection As IJDProfileSection

    If pHelper.ObjectType(oSupport1) = SDOBJECT_STIFFENERSYSTEM Then
            ' Support 1 is a profile....
            Set oProfileObject.object = oSupports(1)
            Set oPlateObject.object = oSupports(2)
            Set oDCrossSection = oProfileObject.object
            bSup1IsProfile = True

    Else
            ' Support 2 is a profile...
            Set oProfileObject.object = oSupports(2)
            Set oPlateObject.object = oSupports(1)
            Set oDCrossSection = oProfileObject.object
            bSup1IsProfile = False

    End If
    
    'Get flange length...
    dProfileWebLen = oProfileObject.WebLength
'    MsgBox "Web Length = " & dProfileWebLen, vbInformation, "For testing only..."
    dProfileFlangeLen = oProfileObject.FlangeLength
    
    'Find out the crossection type of the profile and set parameters accordly...
    Select Case oDCrossSection.CrossSectionType
        Case "UA"
            Select Case dProfileWebLen
                Case 0.1
'                    MsgBox "UA 100x75x7", vbInformation, "Testing..."
                    paHeight = 0.09
                    paWidth = 0.07
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.125, 0.15
'                    MsgBox "UA 125x75x7 or UA 150x90x9 or UA 150x90x12", vbInformation, "Testing..."
                    paHeight = 0.115
                    paWidth = 0.08
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                 Case 0.2
'                    MsgBox "UA 200x90x9", vbInformation, "Testing..."
                    paHeight = 0.17
                    paWidth = 0.11
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.25
'                    MsgBox "UA 250x90x10.5 or UA 250x90x11.5", vbInformation, "Testing..."
                    paHeight = 0.2
                    paWidth = 0.13
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.35
'                    MsgBox "UA 350x100x12/17 or UA 350x100x11/17", vbInformation, "Testing..."
                    paHeight = 0.3
                    paWidth = 0.18
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.3
'                    MsgBox "UA 300x90x...", vbInformation, "Testing..."
                    paHeight = 0.25
                    paWidth = 0.16
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.4
'                    MsgBox "UA 300x90x...", vbInformation, "Testing..."
                    paHeight = 0.35
                    paWidth = 0.21
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case Else
'                   MsgBox "UA Else", vbInformation, "Testing..."
                    paHeight = 0.09
                    paWidth = 0.07
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                    
             End Select
             
        Case "B_TSN", "B"
            Select Case dProfileWebLen
                Case 0.18, 0.2
'                    MsgBox "BPL 180x9.5 or BPL 200x10", vbInformation, "Testing..."
                    paHeight = 0.15
                    paWidth = 0.1
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.23
'                    MsgBox "BPL 230x11", vbInformation, "Testing..."
                    paHeight = 0.17
                    paWidth = 0.11
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case 0.25
'                    MsgBox "BPL 230x11", vbInformation, "Testing..."
                    paHeight = 0.2
                    paWidth = 0.13
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                Case Else
'                   MsgBox "Bulb Else", vbInformation, "Testing..."
                    paHeight = 0.09
                    paWidth = 0.07
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
                
                End Select
                
        Case "BUT"
        
        ' All sizes noted in document do not exist in the current set of BUT's...
            If (dProfileWebLen < 0.425) Then
'                    MsgBox "In less then state...", vbInformation, "For testing only..."
                    paHeight = dProfileWebLen - 0.03
                    paWidth = dProfileWebLen / 2
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            ElseIf (dProfileWebLen < 0.564) Then
                    paHeight = 0.4
                    paWidth = 0.25
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            ElseIf (dProfileWebLen < 0.644) Then
                    paHeight = 0.5
                    paWidth = 0.3
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            ElseIf (dProfileWebLen < 0.719) Then
                    paHeight = 0.58
                    paWidth = 0.35
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            ElseIf (dProfileWebLen < 0.79) Then
                    paHeight = 0.65
                    paWidth = 0.4
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            ElseIf (dProfileWebLen < 0.869) Then
                    paHeight = 0.73
                    paWidth = 0.43
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            ElseIf (dProfileWebLen < 0.924) Then
                    paHeight = 0.8
                    paWidth = 0.47
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
             ElseIf (dProfileWebLen < 1#) Then
                    paHeight = 0.85
                    paWidth = 0.5
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            Else
'                    MsgBox "In BUT Else...", vbInformation, "For testing only..."
                    paHeight = dProfileWebLen - 0.03
                    paWidth = dProfileWebLen / 2
                    paNose1 = 0.015
                    paNose2 = 0.015
                    paScallop = 0.005
            End If
            
        Case Else ' If cross section is not defined in requirments...
        
'            MsgBox "Cross section Else", vbInformation, "Testing..."
            paHeight = 0.09
            paWidth = 0.07
            paNose1 = 0.015
            paNose2 = 0.015
            paScallop = 0.005
            
    End Select

    'Set the sizes based in relationship to the proflie support...
    If bSup1IsProfile Then
        pPRL.Add "Width1", paWidth
'        MsgBox "Width = " & paWidth, vbInformation, "Testing message..."
        pPRL.Add "Height1", paHeight
'        MsgBox "Width = " & paHeight, vbInformation, "Testing message..."
        pPRL.Add "Nose1", paNose1
'        MsgBox "Width = " & paNose1, vbInformation, "Testing message..."
        pPRL.Add "Nose2", paNose2
'        MsgBox "Nose2 = " & paNose2, vbInformation, "Testing message..."
        
    Else
        pPRL.Add "Width1", paHeight
'        MsgBox "Width = " & paHeight, vbInformation, "Testing message..."
        pPRL.Add "Height1", paWidth
'        MsgBox "Height = " & paWidth, vbInformation, "Testing message..."
        pPRL.Add "Nose1", paNose2
'        MsgBox "Nose1 = " & paNose1, vbInformation, "Testing message..."
        pPRL.Add "Nose2", paNose1
'        MsgBox "Nose2 = " & paNose2, vbInformation, "Testing message..."
       
    End If
    
    ' Set the rest of the values...
    pPRL.Add "Scallop", paScallop


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



