VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketFlangeSel"
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

Const m_SelectorProgid As String = "SMBcktFlangeRules.BracketFlangeSel"
Const m_SelectorName   As String = "SMBcktFlangeRules.BracketFlangeSel"
Const m_FamilyProgid   As String = ""

Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketFlangeRules\BracketFlangeSel.cls"
Private METHOD As String

Implements IJDUserSymbolServices
' *******************************************************************************************
Public Sub SelectorInputs( _
    pIH As IJDInputsHelper _
)
    On Error GoTo ErrorHandler
    METHOD = "SelectorInputs"
    
    'define some inputs here
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
' *******************************************************************************************
Public Sub SelectorQuestions( _
    pQH As IJDQuestionsHelper _
)
    On Error GoTo ErrorHandler
    METHOD = "SelectorQuestions"

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
' *******************************************************************************************
Public Sub SelectorLogic( _
    pSL As IJDSelectorLogic _
)
    On Error GoTo ErrorHandler
    METHOD = "SelectorLogic"
    
    ' ---------------------------------------------------------------------------------------
    ' Add a default flange contour. This line can be removed when more appropriate selector
    ' logic is implemented below.
    
    pSL.Add "BP_SM_FlatContour"
    
    ' ---------------------------------------------------------------------------------------
    
    ' Create the utility that will be used to get the bracket plate object from the flange
    ' smart occurrence.
    Dim oStrEntUtils As IJStructEntityUtils
    Set oStrEntUtils = New StructEntityUtils
    
    ' Create the utility that will be used to get the bracket smart occurrence from the
    ' bracket plate object.
    Dim oBracketAttr As IJBracketAttributes
    Set oBracketAttr = New PlateUtils
    
    ' Get the bracket plate object.
    Dim oBracketPlate As IJPlate
    Set oBracketPlate = oStrEntUtils.GetFlangedPlate(pSL.SmartOccurrence)
    
    ' Get the bracket contour smart occurrence.
    Dim oBracketSO As IJSmartOccurrence
    oBracketAttr.GetBracketByPlaneSO oBracketPlate, oBracketSO
    
    If Not oBracketSO Is Nothing And Not oBracketPlate Is Nothing Then
        ' The bracket smart occurrence can now be used to get the selected item
        ' by doing:
        '
        ' oBracketSO.Item
        '
        ' If any information about the bracket plate object is required, the plate
        ' can be queried directly using:
        '
        ' oBracketPlate.xyz
        '
        ' for the required information.
    End If
    
    ' Cleanup.
    Set oStrEntUtils = Nothing
    Set oBracketAttr = Nothing
    Set oBracketPlate = Nothing
    Set oBracketSO = Nothing
    
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
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function
' ********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub
' ********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
' ********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
' ********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
' ********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


