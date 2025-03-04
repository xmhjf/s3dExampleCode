VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HoleRootSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\SmartOccurrence\HoleRules
'
' File: HoleRootSel.cls
'
' Author: Hole Mgmt Team
'
' Abstract: root selector for hole trace smart class
'
' Description: This is the root selector for the hole trace smart class. It
'              will decide between a general hole and a cableway hole based
'              upon the parent outfitting of the hole trace.
'******************************************************************************

Option Explicit

Const m_SelectorProgid As String = "HoleRules.HoleRootSel"
Const m_SelectorName As String = "HoleRules.HoleRootSel"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
  
    'the hole trace is the only input
    pIH.SetInput INPUT_HOLETRACE
  
    Exit Sub

ErrorHandler:
    pIH.ReportError
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
  
    'none defined at this time
  
    Exit Sub

ErrorHandler:
    pQH.ReportError
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
        
    'get the hole trace
    Dim oHoleObject As Object
    Set oHoleObject = pSLH.InputObject(INPUT_HOLETRACE)

    ' ********************* Start Selection Code ****************
    
    'determine if outfitting is a single cableway
    Dim oHoleTrace As IJHoleTraceAE
    Set oHoleTrace = oHoleObject
  
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    
    'single cableway hole goes to the cableway hole selector
    'all others go to general holes
    
    'how many outfitting objects are there
    If oOutfitting.Count = 1 Then
        Dim oObject As Object
        Dim oHMMiddleHelper As IJHMMiddleHelper
        Dim bSupportHoleAspect As Boolean
        
        Set oHMMiddleHelper = New CHMMiddleHelper
        'For Loop input object - this will only go around once
        For Each oObject In oOutfitting
            bSupportHoleAspect = False
            On Error Resume Next
            bSupportHoleAspect = oHMMiddleHelper.DoesObjectSupportHoleAspect(oObject)
            If bSupportHoleAspect = True Then
                pSLH.Add "GeneralHoles"
                Exit For
            Else
                If TypeOf oObject Is IJRteCablewayPathFeat And Not ISImportHoleTrace(oHoleTrace) Then
                    pSLH.Add "CablewayHoles"
                    Exit For
                Else
                    pSLH.Add "GeneralHoles"
                    Exit For
                End If
            End If
        Next oObject
        Set oObject = Nothing
        Set oHMMiddleHelper = Nothing
    Else
        'zero outfitting means stand-alone hole - use the GeneralHoles item
        'more than one outfitting - use the GeneralHoles item
        pSLH.Add "GeneralHoles"
    End If
    
    Set oOutfitting = Nothing
    Set oHoleTrace = Nothing
    
    ' ********************* End Selection Code ******************
    
    Set oHoleObject = Nothing
    Exit Sub

ErrorHandler:
    pSLH.ReportError
End Sub
Private Function ISImportHoleTrace(oHoleTraceAE As IJHoleTraceAE) As Boolean
    Dim oAssoc As IJDAssocRelation
    Dim oTargetCollForTraceComplexString As IJDTargetObjectCol
    Dim ItemMonikerForTraceComplexString As IUnknown
    Dim oTargetCollForTraceShape As IJDTargetObjectCol
    Dim oPOM As IJDPOM
    Dim strClsId As String
    Dim oObject As IJDObject
    
    Set oObject = oHoleTraceAE
    Set oPOM = oObject.ResourceManager
    Set oObject = Nothing
    Set oAssoc = oHoleTraceAE
    Set oTargetCollForTraceComplexString = oAssoc.CollectionRelations("IJHoleTraceAE", "HoleTraceCmplxStr_ORIG")
    Set oTargetCollForTraceShape = oAssoc.CollectionRelations("IJHoleTraceAE", "HoleTraceShape_ORIG")
    Set oAssoc = Nothing
    
    '''''''''''''''''''''For imported trace the relation will be with Complex string only it will not have any relation with trace shape
    
    If Not oTargetCollForTraceComplexString Is Nothing Then
     If oTargetCollForTraceShape.Count = 0 And _
             oTargetCollForTraceComplexString.Count > 0 Then
            'Set ItemMoniker = GetMonikerFromObject(oTargetColl.Item(1))
            Set ItemMonikerForTraceComplexString = oPOM.GetObjectMoniker(oTargetCollForTraceComplexString.Item(1))
            strClsId = oPOM.ClsidOf(ItemMonikerForTraceComplexString)
            If (0 = StrComp(strClsId, "{96EBC514-D39E-11D1-9559-0060973D4824}")) Then   'Classid of Complexstring3d
                ISImportHoleTrace = True
            End If
     End If
    End If
    Set oPOM = Nothing
    Set oTargetCollForTraceComplexString = Nothing
    Set ItemMonikerForTraceComplexString = Nothing
    Set oTargetCollForTraceShape = Nothing

End Function

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
   
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractSelector pSelector
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
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSLH As IJDSelectorLogic
  Set pSLH = New SelectorLogic
  pSLH.Representation = pRep
  SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
 
