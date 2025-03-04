VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2004 Intergraph Corporation. All rights reserved.
'
'File : SplitSel.cls
'
'Author : M. Holderer
'
'Description :
'    root selector for Splitting connections
'
'History:
'   28-jun-2004     moh             created
'*******************************************************************************************************

Option Explicit

Private Const MODULE = "SplitSel"
Private m_oLocalizer As IJLocalizer

Const m_SelectorProgid As String = "SPSSplitMacros.SplitSel"
Const m_SelectorName As String = "SPSSplitMacros.SplitSel"
 
Implements SPSMembers.IJStructILCHelper
Implements SPSMembers.IJStructSOSelectRule
Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo errorHandler
  
  
  Exit Sub
errorHandler:
  pIH.ReportError
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo errorHandler
  ' TODO - Add your question below
  
  Exit Sub
errorHandler:
  pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)

  On Error GoTo errorHandler

'From Patrice,
'        Mike,
'
'        Returning ""  deletes the smart occurrence.
'        Returning "Undefined" deletes the members of the custom assembly.
'        Returning nothing should be an error.
'
'        Patrice

    Dim IHStatus As SPSMembers.StructSOCInputHelperStatus
    Dim ISOC As SPSMembers.IJStructSmartOccurrenceConnection
    Dim iILCHelper As SPSMembers.IJStructILCHelper
    Dim iSOSelectRule As SPSMembers.IJStructSOSelectRule
    Dim strSelection As String

    Set ISOC = pSL.SmartOccurrence
    Set iILCHelper = Me     ' don't care what current def is.  My SelectRule is what gets used.
    Set iSOSelectRule = iILCHelper.SelectionRule(ISOC)
    IHStatus = iSOSelectRule.Evaluate(ISOC, strSelection)

    If IHStatus <> SPSFCInputHelper_Ok Then
        strSelection = ""
    End If

    pSL.Add strSelection

    Exit Sub

errorHandler:
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
'*************************************************************************
'Function
'IJDUserSymbolServices_GetDefinitionName
'
'Abstract
'Used during the execution of IJDDefinitionCollection::GetDefinitionByProgId to get the definition name
'based upon the definitionParameters passed in. It returns the definition name (pDefName) if it already
'exists within the collection. The name of a definition is the identifier of the definition object
'in the definition collection and assures its uniqueness in the given resource manager.
'
'Arguments
'definitionParameters
'
'Return
'
'Exceptions
'
'***************************************************************************

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function
'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Cleans the previous definition up and initializes the new one (Input, Output, Representation,
'RepresenationEvaluation, ...) constructing the symbol definition by (re)defining the inputs, outputs,
'representations, and representation.
'Note:  The previous setting of the definition must be reset in this method before constructing the new definition.
'
'Arguments:
'pDefinition  Symbol definition passed by reference that will be initialized in this method.
'
'Return
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions
'
'***************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
'  pIH.InitAs m_FamilyProgid
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
  
   Dim inputsProp As IMSDescriptionProperties
   inputsProp = pSelector.IJDInputs.Property
   pSelector.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE

End Sub
'*************************************************************************
'Function
'IJDUserSymbolServices_InstanciateDefinition
'
'Abstract
'Instantiates a persistent symbol definition object and initializes it for the first time,
'returning a pointer (ppSymbolDefDisp) to the IDispatch interface of the initialized symbol definition.
'
'Arguments:
'codeBase specifies the URL (or UNC) of the .cab file that can provides the dll associated to the symbol definition object (ActiveX® control packaging).
'definitionParameters  Definition parameters.
'pResourceMgr  resource manager to which the symbol definition will be connected.
'
'Return:
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions:
'
'***************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
'This function is not being called as we call the ISPSFCInputHelper_ExecuteSelectionRule in the IntersectionILCSel class which has the logic of selecting a particular item
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSL As IJDSelectorLogic
  Set pSL = New SelectorLogic
  pSL.Representation = pRep
  SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


'*************************************************************************
'Function
'ISPSFCInputHelper_UserAttributeMgmt
'
'Abstract
'Function on ISPSFACInputHelper to return the UserAttributeMgmt interface
'
'Arguments
'
'Return
'The interface is returned.
'
'Exceptions
'
'***************************************************************************
Private Property Get ISPSILCInputHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
    Set ISPSILCInputHelper_UserAttributeMgmt = Nothing
End Property

Private Function IJStructILCHelper_GetParents(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection, _
        pParent1 As Object, otherParents As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus
    Const MT = "ISPSILCInputHelper_GetParents"
    
    Dim IHStatus As StructSOCInputHelperStatus
    Dim parents As IJElements
    
    Set parents = GetRefCollObjects(ISOC)

    If parents.count < 1 Or parents.count > 2 Then
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_SPLITMACROS_UNEXP_REFCOLLCOUNT, "Split connection is missing a required input. Delete the connection and replace.")
    End If

    Set pParent1 = parents(1)
    
    If parents.count = 2 Then
        Set otherParents = New JObjectCollection
        otherParents.Add parents(2)
    End If

    Set parents = Nothing

    IJStructILCHelper_GetParents = StructSOCInputHelper_Ok

    Exit Function

errorHandler:
    HandleError MODULE, MT
    Err.Clear
    IJStructILCHelper_GetParents = IHStatus
End Function


Private Property Get IJStructILCHelper_SelectionRule(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection) As SPSMembers.IJStructSOSelectRule
    Set IJStructILCHelper_SelectionRule = Me
End Property

Private Function IJStructILCHelper_SetParents(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection, _
    ByVal parent1 As Object, ByVal otherParents As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus

    Const MT = "ISPSILCHelper_SetParents"
    Dim IHStatus As StructSOCInputHelperStatus

    On Error GoTo errorHandler
    
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim parent2 As Object

    If parent1 Is Nothing Then
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_SPLITMACROS_NOPARENT1, "Split connection is missing the member system to split. Delete the connection and replace.")
    End If

    If Not otherParents Is Nothing Then
        If Not TypeOf otherParents Is IJElements Then
            IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_SPLITMACROS_UNKNOWNTYPE_PARENT2, "Object that identifies the split location is not a required linear object. Delete the connection and replace.")
        ElseIf otherParents.count > 1 Then
            IHStatus = StructSOCInputHelper_BadNumberOfObjects
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_SPLITMACROS_NOSECOND_PARENT, "Split connection is missing the object that identifies the split location. Delete the connection and replace.")
        ElseIf otherParents.count = 1 Then
            Set parent2 = otherParents(1)
        End If
    End If

    If parent1 Is parent2 Then
        IHStatus = StructSOCInputHelper_DuplicateObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_SPLITMACROS_SCNDPARENT_SAMEAS_FSTPARENT, "Split connection is using the same member for both the member to split and to define the split location. Delete the connection and replace.")
    End If
    
    'set interfaces for supported member
    Set oRC = GetRefColl(ISOC)
    oRC.IJDEditJDArgument.RemoveAll

    oRC.IJDEditJDArgument.SetEntity 1, parent1, ConstIJGeometry, "SplitRefColl_DEST"
    If Not parent2 Is Nothing Then
        oRC.IJDEditJDArgument.SetEntity 2, parent2, ConstIJGeometry, "SplitRefColl_DEST"
    End If
    
    IJStructILCHelper_SetParents = StructSOCInputHelper_Ok

    Exit Function

errorHandler:
    HandleError MODULE, MT
    Err.Clear
    IJStructILCHelper_SetParents = IHStatus

End Function

Private Function IJStructILCHelper_UpdateValuesFromPosition(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection) As SPSMembers.StructSOCInputHelperStatus
    IJStructILCHelper_UpdateValuesFromPosition = StructSOCInputHelper_Ok
End Function

Private Property Get IJStructILCHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
    Set IJStructILCHelper_UserAttributeMgmt = Nothing
End Property

Private Function IJStructILCHelper_ValidateParents(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection, _
    ByVal inputParent1 As Object, ByVal inputOtherParents As SPSMembers.IJElements, _
    pRelationParent1 As Object, pRelationOtherParents As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus

    
    Const MT = "ISPSILCHelper_ValidateParents"
    Dim IHStatus As StructSOCInputHelperStatus
    Dim parent1 As Object, parent2 As Object
    Dim splitStat As splitStatus
    Dim ok As Boolean
    Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double
    
    On Error GoTo errorHandler
    
    ' check type
    ' intersect the parents
    
    IHStatus = StructSOCInputHelper_UnexpectedError
    
    If inputParent1 Is Nothing Then
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If

    If Not inputOtherParents Is Nothing Then
        If Not TypeOf inputOtherParents Is IJElements Then
            IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            GoTo wrapup
        ElseIf inputOtherParents.count > 1 Then
            IHStatus = StructSOCInputHelper_BadNumberOfObjects
            GoTo wrapup
        ElseIf inputOtherParents.count = 1 Then
            Set parent2 = inputOtherParents(1)
        End If
    End If
    
    If Not parent2 Is Nothing Then
        If TypeOf inputParent1 Is IJStructSystemParent Then
            Set parent1 = inputParent1
            If TypeOf parent2 Is IJStructSystemParent Then
                ' ssSplitNone tolerates positions at the ends of member systems.
                ' so, here we don't want to block something that the specific def will check for.
                splitStat = ssSplitNone
            ElseIf ObjectIsOkSurface(parent2) Then
                splitStat = ssSplitFirst
            Else
                IHStatus = StructSOCInputHelper_InvalidTypeOfObject
                GoTo wrapup
            End If
    
        ElseIf ObjectIsOkSurface(inputParent1) Then
            Set parent1 = inputParent1
            If Not TypeOf parent2 Is IJStructSystemParent Then
                IHStatus = StructSOCInputHelper_InvalidTypeOfObject
                GoTo wrapup
            End If
            splitStat = ssSplitSecond
    
        Else
            IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            GoTo wrapup
        End If
    
        ComputeIntersection parent1, parent2, splitStat, 0#, x1, y1, z1, x2, y2, z2, ok
        
        If Not ok Then
            IHStatus = StructSOCInputHelper_BadGeometry
        ElseIf RedundantConnectionExists(ISOC, parent1, parent2) Then
            IHStatus = StructSOCInputHelper_DuplicateObject
        Else
            Set pRelationParent1 = parent1
            Set pRelationOtherParents = New JObjectCollection
            pRelationOtherParents.Add parent2
            IHStatus = StructSOCInputHelper_Ok
        End If
    
    ' else parent2 is Nothing.   Just check that parent1 is okay.
    ElseIf Not TypeOf inputParent1 Is IJStructSystemParent Then
        IHStatus = StructSOCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    
    Else
        Set pRelationParent1 = inputParent1
        IHStatus = StructSOCInputHelper_Ok
    End If

wrapup:
    IJStructILCHelper_ValidateParents = IHStatus
    Exit Function

errorHandler:
    Exit Function
End Function

Private Function IJStructSOSelectRule_Evaluate(ByVal iSmartOccurrence As Object, strSelectedItem As String) As SPSMembers.StructSOCInputHelperStatus

    On Error GoTo errorHandler
    Dim socStatus As SPSMembers.StructSOCInputHelperStatus
    
    socStatus = EvaluateSplitRule(iSmartOccurrence, True, strSelectedItem)

    IJStructSOSelectRule_Evaluate = socStatus
    Exit Function

errorHandler:
    IJStructSOSelectRule_Evaluate = StructSOCInputHelper_UnexpectedError
    Err.Raise E_FAIL
End Function
  
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
