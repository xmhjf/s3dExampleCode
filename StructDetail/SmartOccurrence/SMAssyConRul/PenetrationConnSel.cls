VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PenetrationConnSel"
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
'
' Modify class name to particular rule
Private Const m_sClassName As String = "PenetrationConnSel"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgid As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Setting penetrated and penetrating objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", strError).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
          'Define allowed answers to questions
'        Dim codeCol As New Collection
'        codeCol.Add "High"
'        codeCol.Add "Medium"
'        codeCol.Add "Low"
'        pQH.DefineCodeList "StressLevelCol", codeCol
        
        pQH.SetQuestion "StressLevel", gsHigh, "StressLevelCol"
        
      'Define allowed answers to questions
'        Dim codeCol2 As New Collection
'        codeCol2.Add gsDrop
'        codeCol2.Add gsSlide
'        pQH.DefineCodeList "AssyMethodCol", codeCol2
          
        'Define questions
        pQH.SetQuestion "AssyMethod", _
                        gsDrop, _
                        "AssyMethodCol", _
                        "SetAnswerToAssyMethod", _
                        m_sRuleProgID
                                
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler

    Dim strError As String
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    
    strError = "Retrieving assembly connection from Smart Occurrence."
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence

    strError = "Retrieving assembly connection penetrating object type."
    Select Case oAssyConn.ConnectedObject2Type  'Penetrating object
        Case SDOBJECT_PLATE, SDOBJECT_STIFFENER, SDOBJECT_BEAM, SDOBJECT_SEAM
                    pSL.Add "StructuralPenetrations"
        Case SDOBJECT_PIPE, SDOBJECT_VENT, SDOBJECT_CABLE, SDOBJECT_CABLEWAY, SDOBJECT_EQUIPMENT
                    pSL.Add "OutfittingPenetrations"
        Case Else
            strError = "Invalid assembly connection penetrating object type found."
            GoTo ErrorHandler
    End Select
  
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
    IJDUserSymbolServices_GetDefinitionName = m_sRuleName
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
    pIH.Definition = pSelector
    pIH.InitAs m_sFamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_sRuleProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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

' This method sets answer to "AssyMethod" question based on planning assembly info.
Public Sub SetAnswerToAssyMethod( _
                  ByVal oInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                  ByRef oArgument As Object)
    On Error GoTo ErrorHandler
    
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymDef As IMSSymbolEntities.IJDSymbolDefinition
    Dim oSL As IJDSelectorLogic
    
    Set oInputDG = oInput
    Set oSymDef = oInputDG.Definition
    
    Set oSL = New SelectorLogic
    oSL.Representation = oSymDef.IJDRepresentations(1)
    
    Dim oACWrapper As New StructDetailObjects.AssemblyConn
    Dim oPenetrated As Object
    Dim oPenetrating As Object
    
    Set oACWrapper.object = oSL.SmartOccurrence
    Set oPenetrated = oACWrapper.ConnectedObject1
    Set oPenetrating = oACWrapper.ConnectedObject2
        
    Dim oSDOHelper As New StructDetailObjects.Helper
    Dim oAssembly As IJAssembly
    Dim sBuildMethod As String
    
    sBuildMethod = gsDrop
    
    On Error Resume Next
    Set oAssembly = oSDOHelper.FirstMeet(oPenetrated, oPenetrating)
    
    Err.Clear
    On Error GoTo ErrorHandler
    
    If Not oAssembly Is Nothing Then
        '
        ' Call PlanningObjects.PlnAssembly.BuildMethod to get build method.
        ' PlanningObjects is in Middle, while this probject is in Content which
        ' is built after PlanningObjects during Smart 3D internal build process.
        
        Dim oPlnAssemblyHelper  As PlanningObjects.PlnAssembly
        Set oPlnAssemblyHelper = New PlanningObjects.PlnAssembly
        
        Set oPlnAssemblyHelper.object = oAssembly
        sBuildMethod = oPlnAssemblyHelper.BuildMethod(oPenetrated)
        
        If sBuildMethod = vbNullString Then
            sBuildMethod = gsDefaultValue
        End If
    End If
    
    oSL.Answer("AssyMethod") = sBuildMethod
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetAnswerToAssyMethod").Number

End Sub


