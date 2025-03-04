VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FittedAsmConnSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : FittedAsmConnSel.cls
'
'Author : R. Peringoth
'
'Description :
'    Macro for selecting a fitted connection based on input properties
'
'History:
'
'
'********************************************************************

Option Explicit

Const m_SelectorProgid As String = "SPSACMacros.FittedAsmConnSel"
Const m_SelectorName As String = "SPSACMacros.FittedAsmConnSel"

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
    Dim oSupped1Port As ISPSSplitAxisPort, oSupped2Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic, oSupped2Part As ISPSMemberPartPrismatic
    Dim oCatalogPOM As IJDPOM
    Dim oNamingContext As IJDNamingContextObject
    Dim oItem As Object
    
    Set oStructConn = pSL.SmartOccurrence
    oStructConn.enumPorts colPorts

    If colPorts.count = 2 Then
        Set oSupped1Port = colPorts.Item(1)
        Set oSupped2Port = colPorts.Item(2)
        Set oSupped1Part = oSupped1Port.Part
        Set oSupped2Part = oSupped2Port.Part
            
        If (oSupped1Part.CrossSection.SectionType = "L") And (oSupped2Part.CrossSection.SectionType = "L") Then
            Set oCatalogPOM = GetCatalogResourceManager()
            Set oNamingContext = New NamingContextObject
            ' added this in v5 to pick fitted with 0 clearance for angles
            On Error Resume Next
            Set oItem = oNamingContext.ObjectMoniker(oCatalogPOM, "FittedAsmConn_No_Clearance")
            On Error GoTo ErrorHandler
            If Not oItem Is Nothing Then
                pSL.Add "FittedAsmConn_No_Clearance"
            Else
                'it may not exist in older catalogs so pick "FittedAsmConn_1" as before
                pSL.Add "FittedAsmConn_1"
            End If
        Else
            pSL.Add "FittedAsmConn_1"
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







