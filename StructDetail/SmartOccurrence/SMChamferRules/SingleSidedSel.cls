VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleSidedSel"
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

Const m_SelectorProgid As String = CUSTOMERID + "ChamferRules.SingleSidedSel"
Const m_SelectorName As String = CUSTOMERID + "ChamferRules.SingleSidedSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "ChamferRules\SingleSidedSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add input description below
    
    ' Part to be chamfered
    ' Opposite Part - determines if a chamfer is needed and if so what size
    
    pIH.SetInput CHAMFER_PART
    pIH.SetInput OPPOSITE_PART
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    ' TODO - Add the code of selector logic below
    
    Dim sChamferType As String
    sChamferType = pSLH.SelectorAnswer(CUSTOMERID + "ChamferRules.RootChamferSel", gsChamferType)
    Dim sChamferWeld As String
    sChamferWeld = pSLH.SelectorAnswer(CUSTOMERID + "ChamferRules.RootChamferSel", gsChamferWeld)
    
    Dim oParentSmartOccurence As IJSmartOccurrence
    
    'Get Chamfer Parent Smart Occurance
    GetSmartOccurrenceParent pSLH.SmartOccurrence, oParentSmartOccurence
    
    'Select Chamfer Type
    If TypeOf oParentSmartOccurence Is IJStructAssemblyConnection Then
        If sChamferType = "Obj1Offset" Then
            pSLH.Add "SingleSidedOffset"
        Else
            pSLH.Add "SingleSidedBase"
        End If
    ElseIf TypeOf oParentSmartOccurence Is IJStructFeature Then
        Dim oWebCut As IJStructFeature
        Set oWebCut = oParentSmartOccurence
        If (oWebCut.get_StructFeatureType = SF_WebCut) Then
            If sChamferType = "Obj1Offset" Or sChamferType = "Obj2Offset" Then
                pSLH.Add "SingleSidedOffset"
            ElseIf sChamferType = "Obj1Base" Or sChamferType = "Obj2Base" Then
                pSLH.Add "SingleSidedBase"
            End If
        End If
    Else
        'Select Chamfer
        If sChamferType = "Obj1Offset" Then
            pSLH.Add "SingleSidedOffsetPC"
        ElseIf sChamferType = "Obj2Offset" Then
            pSLH.Add "SingleSidedOffsetPC"
        ElseIf sChamferType = "Obj1Base" Then
            pSLH.Add "SingleSidedBasePC"
        ElseIf sChamferType = "Obj2Base" Then
            pSLH.Add "SingleSidedBasePC"
        ElseIf sChamferType = "Obj1BaseObj2Offset" Then
            If sChamferWeld = "First" Then
                pSLH.Add "SingleSidedBasePC"
            Else
                pSLH.Add "SingleSidedOffset"
            End If
        ElseIf sChamferType = "Obj1OffsetObj2Base" Then
            If sChamferWeld = "First" Then
                pSLH.Add "SingleSidedOffsetPC"
            Else
                pSLH.Add "SingleSidedBase"
            End If
        End If
    End If
 Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
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
    pIH.definition = pSelector
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

