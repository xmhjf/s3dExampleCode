VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndByPlateSel"
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

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.StiffEndByPlateSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.StiffEndByPlateSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffEndByPlateSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", strError).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Define questions
    strError = "Setting questions."
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList", "SetEndCutType", CUSTOMERID + "AssyConRul.AssyConnSelCM"
    
    pQH.SetQuestion "PlaceBracket", "No", "BooleanCol", "PlaceBracketCM", CUSTOMERID + "AssyConRul.StiffEndByPlateSel"
    
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", strError).Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim bChamferPlateEdge As Boolean
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim iPort2TopoContext As IMSStructConnection.eUSER_CTX_FLAGS  'Bounding object port
    Dim iPort2CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim iPlatePortContext As enmPortBasicContext
    
    strError = "Retrieving assembly connection from Smart Occurrence."
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence

    'Determine which definition rule to fire based on port context of bounding plate.
    strError = "Retrieving assembly connection bounding port info."
    Call oAssyConn.Port2Topology(iPort2TopoContext, iPort2CrossSectEntity)  'Plate is port 2
    iPlatePortContext = PortBasicContext(iPort2TopoContext)

    Select Case iPlatePortContext
        Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
            pSL.Add "StiffenerEndToPlateFace"
        Case PORT_BASIC_CONTEXT_LATERAL

            bChamferPlateEdge = AC_PlateEdgeChamferedByStiffenerEdge(oAssyConn.object)
            If bChamferPlateEdge Then
                pSL.Add "StiffenerEndToChamferEdge"
            End If
            
            pSL.Add "StiffenerEndToPlateEdge"
        Case Else
            strError = "Unknown value: iPlatePortContext=" & iPlatePortContext
            GoTo ErrorHandler
    End Select
    
    Set oAssyConn = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
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
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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
Public Sub PlaceBracketCM(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    'get the answer to the EndCutType question
    Dim sEndCutType As String
    sEndCutType = pSL.SelectorAnswer(CUSTOMERID + "AssyConRul.StiffEndByPlateSel", "EndCutType")
    
    If sEndCutType = "Bracketed" Then
        pSL.Answer("PlaceBracket") = "Yes"
    Else
        pSL.Answer("PlaceBracket") = "No"
    End If
        

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "PlaceBracketCM").Number
End Sub
