VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCIRCLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 '******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    "SketchCSSymbol.CCIRCLE"
'
'Revision
'
'Description
'    Definition of CIRCLE
'Notes
'
'History
'ARV        11th Feb 02           Creation
'ARV      10th April 02         Added CacheOption in
'                                    IJDUserSymbolServices_InstanciateDefinition
'                                    for solving Caching Problem
'ARV      19th June 2002        Fix for TR 28200
'ARV       21th June 2002        implemented method IsValidParameters
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'ARN       19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const m_progID = "SketchCSSymbol.CCircle"
Const EPSILON = 0.000001
Private Const E_FAIL = -2147467259
Private Const INPUT_1 = "R"
Private Const INPUT_2 = "Cardinality"
Private Const INPUT_3 = "Xref"
Private Const INPUT_4 = "Yref"

Private Const OUTPUT_1 = "Circle"

Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_EditOccurence = False
    Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
      
    SymbolDefinition pSymbolDefinition
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_GetDefinitionName = m_progID
    Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As IMSSymbolEntities.DSymbolEntitiesFactory
    Set oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory

    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition

    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_progID
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
'Arnab        19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'    oSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    Exit Function

ErrorHandler:
    Err.Raise E_FAIL
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayofinputs())
    On Error GoTo ErrorHandler
    Set m_outputColl = outputcoll
    
    If StrComp(repName, "Symbolic") = 0 Then
        Call Symbolic(arrayofinputs)
    End If
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub
Private Sub Symbolic(ByRef arrayofinputs())
On Error GoTo ErrHandler
    Dim lCardinality        As Long
    Dim XCord               As Double
    Dim YCord               As Double
    Dim dRadius             As Double
    Dim xRef                As Double
    Dim yRef                As Double
    
    'Check for the validity of inputs
    If IsValidParameters(arrayofinputs) = False Then Exit Sub
    
    dRadius = arrayofinputs(1)
    lCardinality = arrayofinputs(2)
    xRef = arrayofinputs(3)
    yRef = arrayofinputs(4)

    Select Case lCardinality
    
    Case 2: XCord = 0#
            YCord = -dRadius
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 3: XCord = 0#
            YCord = dRadius
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 4: XCord = dRadius
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 5: XCord = -dRadius
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
    
    Case 100: XCord = xRef - dRadius 'User defined
            YCord = yRef '- dRadius
            TranslateSymbol arrayofinputs, XCord, YCord
    
    Case Else: XCord = 0#
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
    End Select
    Exit Sub
ErrHandler:
    Error.Raise E_FAIL
End Sub
'
' SAMPLE CODE TO CREATE A CIRCLE SYMBOL
Private Function Graphics(ByRef arrayofinputs()) As Circle3d
  
    On Error GoTo ErrHandler
    Dim oGeomFactory        As New GeometryFactory
    Dim oCircle             As Circle3d
    Dim dRadius             As Double

    dRadius = arrayofinputs(1)
    
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(m_outputColl.ResourceManager, 0#, 0#, 0#, _
                    0#, 0#, 1#, dRadius)
   
    Set Graphics = oCircle
    
    Set oCircle = Nothing
    
    Set oGeomFactory = Nothing

    Exit Function
    
ErrHandler:
    Error.Raise E_FAIL
End Function

Private Sub TranslateSymbol(ByRef arrayofinputs(), XTrans As Double, YTrans As Double)
On Error GoTo ErrHandler
    Dim oTrans4x4           As IJDT4x4
    Dim oVector             As IJDVector

    Dim oCircle             As Circle3d
    
    Set oTrans4x4 = New DT4x4
    Set oVector = New DVector
    
    oVector.Set XTrans, YTrans, 0#
    oTrans4x4.LoadIdentity
    oTrans4x4.Translate oVector
    
    Set oCircle = Graphics(arrayofinputs)
    
    ' if symbol is not there just exit
    If oCircle Is Nothing Then Exit Sub
    
    oCircle.Transform oTrans4x4

    m_outputColl.AddOutput OUTPUT_1, oCircle
    
    Set oCircle = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    Exit Sub
    
ErrHandler:
    Error.Raise E_FAIL
End Sub



Sub SymbolDefinition(ByRef pSymbolDefinition As Object)
  
    On Error GoTo ErrorHandler
    
    ' Create a new input by new operator
    Dim oInput(1 To 4) As IMSSymbolEntities.IJDInput
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Dim iNumInput As Integer
    Dim iNumOutput As Integer
    Dim iNumRep As Integer
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
  
    ' Create a defaultValue
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    PC.Type = igValue
    
    For iNumInput = 1 To 4
      Set oInput(iNumInput) = New IMSSymbolEntities.DInput
    Next
      
    oInput(1).Name = INPUT_1
    oInput(1).Description = "Radius"
    oInput(1).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1#
    oInput(1).DefaultParameterValue = PC

    oInput(2).Name = INPUT_2
    oInput(2).Description = "Cardinality"
    oInput(2).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1
    oInput(2).DefaultParameterValue = PC
    
    oInput(3).Name = INPUT_3
    oInput(3).Description = "Xref"
    oInput(3).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    oInput(3).DefaultParameterValue = PC
    
    oInput(4).Name = INPUT_4
    oInput(4).Description = "Yref"
    oInput(4).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    oInput(4).DefaultParameterValue = PC
    
    ' Set the input to the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    For iNumInput = 1 To 4
      oInputs.SetInput oInput(iNumInput), iNumInput
    Next

    ' Create the output
    Dim oOutput(1 To 1) As IMSSymbolEntities.IJDOutput
    
    For iNumOutput = 1 To 1
      Set oOutput(iNumOutput) = New IMSSymbolEntities.DOutput
    Next
    
    oOutput(1).Name = OUTPUT_1
    oOutput(1).Description = "Description for OUTPUT_1"
    oOutput(1).Properties = igDESCRIPTION_INIT

  ' Define the representation "Symbolic"
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation

    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = pSymbolDefinition
    
    Set oOutputs = oRep
    
    oRep.Name = "Symbolic"
    oRep.Description = "Description for Symbolic"
    oRep.Properties = igREPRESENTATION_ISVBFUNCTION
    oRep.RepresentationId = 1


    oOutputs.SetOutput oOutput(1)
    oReps.SetRepresentation oRep

    Dim oVbFuncRep(1 To 1) As IJDRepresentationEvaluation
    
    For iNumRep = 1 To 1
      Set oVbFuncRep(iNumRep) = New DRepresentationEvaluation
      oVbFuncRep(iNumRep).Properties = igREPRESENTATION_HIDDEN
      oVbFuncRep(iNumRep).Type = igREPRESENTATION_VBFUNCTION
      oVbFuncRep(iNumRep).ProgId = m_progID
    Next iNumRep
    
    Dim oScripts As IMSSymbolEntities.IJDRepresentationEvaluations
    Set oScripts = pSymbolDefinition
    
    oVbFuncRep(1).Name = "Symbolic"
    oVbFuncRep(1).Description = "script for the Symbolic representation"
    oScripts.AddRepresentationEvaluation oVbFuncRep(1)

    Exit Sub
ErrorHandler:
    Error.Raise E_FAIL
End Sub


Private Function IsValidParameters(ByRef arrayofinputs()) As Boolean
On Error GoTo ErrorHandler
    Dim oErrors As IJEditErrors
    Dim dRadius As Double
    Dim lCardinality As Long
    
    Set oErrors = New JServerErrors
    dRadius = arrayofinputs(1)
    lCardinality = arrayofinputs(2)
    
    IsValidParameters = True
  ' Fix For TR 28200
    'If Symbolattribute value is negative then give message to the user
    If dRadius < EPSILON Or lCardinality < 1 Then
         IsValidParameters = False
         oErrors.Add E_FAIL, "SketchCSSymbol.CCIRCLE", "CrossSection Dimensions are not valid", "Circle"
         GoTo ErrorHandler
    End If

  Exit Function
  
ErrorHandler:
    Error.Raise E_FAIL
End Function
