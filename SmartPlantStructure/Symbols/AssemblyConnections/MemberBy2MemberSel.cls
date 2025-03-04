VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberBy2MemberSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : MemberBy2MemberSel.cls
'
'Author : RP
'
'Description :
'    Macro for selecting a 3 member  connection based on input properties
'
'History:
'
' 04 -14-04  RP   Added check for proper orientation before selecting CornerGussetAsmConn
' 06/13/06  RP   Changes due to impact from curved members(DI#84001)
' 12/11/06  RP   Do not select cornergussetplate  connection when supported member has I-shaped
'                section
'********************************************************************

Option Explicit

Const m_SelectorProgid As String = "SPSACMacros.MemberBy2MemberSel"
Const m_SelectorName As String = "SPSACMacros.MemberBy2MemberSel"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add input description below
    
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
    
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSupping1Port As ISPSSplitAxisPort
    Dim oSupping2Port As ISPSSplitAxisPort
    Dim oSupping1Part As ISPSMemberPartPrismatic
    Dim oSupping2Part As ISPSMemberPartPrismatic
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim strSectionType As String

    
    
    Set oStructConn = pSL.SmartOccurrence
    oStructConn.enumPorts colPorts
    GetCornerGussetInputPorts colPorts, oSuppedPort, oSupping1Port, oSupping2Port
    Set oSuppedPart = oSuppedPort.Part
    Set oSupping1Part = oSupping1Port.Part
    Set oSupping2Part = oSupping2Port.Part
    

    strSectionType = oSuppedPart.CrossSection.SectionType
    If strSectionType <> "W" And strSectionType <> "S" And strSectionType <> "HP" And strSectionType <> "M" Then
        If CheckForCornerBraceAsmConn(oSuppedPart, oSupping1Part, oSupping2Part) Then
            pSL.Add "CornerGussetAsmConn"
'TR139884: Log an error when members are not in the same plane (Bad computation)
        Else
            SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_MEMBS_NOTIN_SAMEPLANE, oStructConn, Nothing
            Err.Raise E_FAIL
        End If
    End If
    
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








