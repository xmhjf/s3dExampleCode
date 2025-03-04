VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitByPenetrationSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************
'
'Copyright (C) 2012 Intergraph Corporation. All rights reserved.
'
'File : SplitByPenetrationSel.cls
'
'Author : Ramesh Kandukuri
'
'Description :
'    Smart marine structure split penetrate connection selection rule.
'
'History:
'
'
'********************************************************************
Const m_SelectorProgid As String = "SMSplitMacros.SplitByPenetrationSel"
Const m_SelectorName As String = "SMSplitMacros.SplitByPenetrationSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "\SharedContent\Src\ShipStructure\SMSplitMacros\SplitByPenetrationSel.cls"
Private Const SI_SplitByPenetration = "SplitByPenetration-1"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo errorHandler

    pIH.SetInput "SplitParentsRefColl"
  
    Exit Sub
    
errorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo errorHandler
    
    Exit Sub
    
errorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub
Public Sub SelectorLogic(pSL As IJDSelectorLogic)

  On Error GoTo errorHandler
  
    pSL.Add SI_SplitByPenetration

    Exit Sub

errorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub

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

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Private Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function
