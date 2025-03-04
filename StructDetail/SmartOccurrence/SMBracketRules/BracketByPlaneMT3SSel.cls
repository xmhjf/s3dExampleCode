VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketByPlaneMT3SSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgid As String = m_sProjectName + ".BracketByPlaneMT3SSel"
Const m_SelectorName As String = m_sProjectName + ".BracketByPlaneMT3SSel"
Const m_FamilyProgid As String = ""

Const QUESTION_BracketByPlaneType As String = "BracketByPlaneType"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "BracketRules\BracketByPlaneMT3SSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' inputs:
    '   Plane that bracket is on
    '   Bracket Plate System
    '   U Point
    '   V Point
    '   Group of supports
    
    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
    pIH.SetInput INPUT_UPOINT
    pIH.SetInput INPUT_VPOINT
    pIH.SetInput INPUT_SUPPORTS
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    pQH.SetQuestion QUESTION_BracketByPlaneType, _
                    "3SLinearTrim", _
                     CUSTOMERID + "BracketByPlane3SCol", _
                    "BracketType"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sBracketByPlaneType As String
    sBracketByPlaneType = pSL.Answer(QUESTION_BracketByPlaneType)
    
'    'get the number of supports selected
'    Dim nSupports As Long
'    Dim oSupportsColl As IJElements
'
'    ' Get Bracket-by-Plane Smart Occurrence:
'    Dim oBracketByPlaneSO As IJSmartOccurrence
'    Set oBracketByPlaneSO = pSL.SmartOccurrence
'
'    ' Create helper for final construction of bracket:
'    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
'    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
'
'    'Get the supports from BracketContourSO.
'    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupportsColl, nSupports
'
'    'Get Supports
'    Dim oSupport3 As Object
'
'    If Not oSupportsColl Is Nothing Then
'        If oSupportsColl.Count > 3 Then
'            Set oSupport3 = oSupportsColl.Item(3)
'        Else
'            Exit Sub
'        End If
'    Else
'        Exit Sub
'    End If
'
'    'Check support  4 and 5 is plate or profile
'    Dim pHelper As StructDetailObjects.Helper
'    Set pHelper = New StructDetailObjects.Helper
'    Dim oProfileSup3 As StructDetailObjects.ProfileSystem
'
'    Select Case pHelper.ObjectType(oSupport3)
'            Case SDOBJECT_PLATESYSTEM
'                    'Get the profile for support four...
'                    Set oProfileSup3 = New StructDetailObjects.ProfileSystem
'                    Set oProfileSup3.object = oSupportsColl.Item(4)
'
'                    'find out if one if one of the profiles is a FB...
'                    If oProfileSup3.SectionType = "FB" Then
'                    Select Case sBracketByPlaneType
'                            Case "3SLinearTrim"
'                                pSL.Add "BBP3SLTSB H1xH2xNxSB"
'                                pSL.Add "BBP3SLTO H1xH2xN"
'
'                            Case "3SCurvedTrim"
'                                pSL.Add "BBP3SCTSB H1xH2xNxSB"
'                                pSL.Add "BBP3SCTO H1xH2xN"
'
'
'
'
'
'
'            End Select
            
    
    Select Case sBracketByPlaneType
        Case "3SLinearTrim"
            pSL.Add "BBP3SLTSB H1xH2xNxSB"
            pSL.Add "BBP3SLTO H1xH2xN"

        Case "3SCurvedTrim"
            pSL.Add "BBP3SCTSB H1xH2xNxSB"
            pSL.Add "BBP3SCTO H1xH2xN"
            
        Case Else
            pSL.Add "BBP3SLTSB H1xH2xNxSB"
            pSL.Add "BBP3SLTO H1xH2xN"
            pSL.Add "BBP3SCTSB H1xH2xNxSB"
            pSL.Add "BBP3SCTO H1xH2xN"
            
    End Select
                
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub
Public Sub BracketType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.Definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
        
    pSL.Answer(QUESTION_BracketByPlaneType) = "3SLinearTrim"
    
              
    Exit Sub
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
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



