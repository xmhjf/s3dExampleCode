VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericRectPlateSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------
'    Copyright (C) 2003 - 2004 Intergraph Corporation. All rights reserved.
'
'
'
'   Selector:
'   ---------
'   It selects the list of avaible Items or SmartClasses in the context of
'   the smart occurrence.
'
'
'History
'    Raj Peringoth      03/26/03      Creation
'---------------------------------------------------------------------------------------


Option Explicit

Const m_SelectorProgid As String = "SPSPartMacros.GenericRectPlateSel"
Const m_SelectorName As String = "SPSPartMacros.GenericRectPlateSel"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  ' TODO - Add your question below
  
  Exit Sub
ErrorHandler:
  pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  ' TODO - Add the code of selector logic below
  
  pSL.Add "GenericRectPlatePart_1"
   
 
  
  Exit Sub
ErrorHandler:
  pSL.ReportError
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
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pSelector
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





