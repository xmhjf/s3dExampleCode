VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisByFCAndOSNoEdgeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrAxisByFCAndOSNoEdgeSel.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'   25Aug2011   - Added new methods to support Inset brace (B-52232)
'    15Sep2011  -   pnalugol
'               Modified/Added methods to support Braces on Flange penetrated cases
'*********************************************************************************************

Const m_sClassName As String = "MbrAxisByFCAndOSNoEdgeSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    ' TODO - Add your question below
    '===============================
    pQH.SetQuestion "TopFlangeOr(Left)Web", gsStraight, gsEndCutShapeCol, "CMSetTopFlangeOrLeftWebShapeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "BottomFlangeOrRightWeb", gsStraight, gsEndCutShapeCol, "CMSetBottomFlangeOrRightWebShapeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "ShapeOutside", gsStraight, gsShapeOutsideCol, "CMSetShapeOutsideAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "TopShapeAtFace", gsNone, gsShapeAtFaceCol, "CMSetTopShapeAtFaceAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "BottomShapeAtFace", gsNone, gsShapeAtFaceCol, "CMSetBottomShapeAtFaceAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
   'questions added to create Top & Bottom Brace
    pQH.SetQuestion "TopBraceType", "Not Applicable", gsBraceTypeCol, "CMSetTopBraceTypeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "TopBraceBoundedEndType", "Structure", gsBraceConnTypeCol, "CMSetTopBraceBoundedEndTypeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "TopBraceBoundedEndValue", "1", , "CMSetTopBraceBoundedEndValueAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "TopBraceBoundingEndType", "Structure", gsBraceConnTypeCol, "CMSetTopBraceBoundingEndTypeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "TopBraceBoundingEndValue", "1", , "CMSetTopBraceBoundingEndValueAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    
    pQH.SetQuestion "BottomBraceType", "Not Applicable", gsBraceTypeCol, "CMSetBottomBraceTypeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "BottomBraceBoundedEndType", "Structure", gsBraceConnTypeCol, "CMSetBottomBraceBoundedEndTypeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "BottomBraceBoundedEndValue", "1", , "CMSetBottomBraceBoundedEndValueAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "BottomBraceBoundingEndType", "Structure", gsBraceConnTypeCol, "CMSetBottomBraceBoundingEndTypeAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    pQH.SetQuestion "BottomBraceBoundingEndValue", "1", , "CMSetBottomBraceBoundingEndValueAnswer", CUSTOMERID + "MbrAC.AxisSelCM"
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of Member Assembly Connection here
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Selecting valid Smart item"
    
    pSelectorLogic.Add "MbrAxis_ToFaceAndOutSideNoEdge"
     
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub
  
'** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
'** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************

'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)

    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
    
End Sub

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

