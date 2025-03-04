VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PadPlateSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------
'    Copyright (C) 2006 Intergraph Corporation. All rights reserved.
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
'    AS      April 21, 2006      Creation
'---------------------------------------------------------------------------------------


Option Explicit

Const m_SelectorProgid As String = "SPSPartMacros.PadPlateSel"
Const m_SelectorName As String = "SPSPartMacros.PadPlateSel"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  pIH.SetInput "MemberPart"
  pIH.SetInput "Surface"
  
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
    Dim oStructPlate As IJStructPlate
    Dim oOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic

    Dim oXSection As IJCrossSection
    
    Set oStructPlate = pSL.SmartOccurrence
    
    Set oOcc = oStructPlate
    Set oRefColl = GetRefCollFromSmartOccurrence(oOcc)
    'get input ports

    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    
    Set oXSection = oSuppedPart.CrossSection.Definition
    
    'TODO: Change the pad type to correct type if content for circular and triangular pads is added
    Select Case oXSection.Type
        Case "HSSC", "CS", "PIPE"
          pSL.Add "PadPlateRectangular"     '"PadPlateCircular"
        Case "L"
          pSL.Add "PadPlateRectangular"     '"PadPlateTriangular"
        Case Else
          pSL.Add "PadPlateRectangular"
    End Select
   ' MsgBox "In PadPlate SelectorLogic"
  
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







