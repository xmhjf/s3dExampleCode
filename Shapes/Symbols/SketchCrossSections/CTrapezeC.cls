VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrapezeC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    "SketchCSSymbol.CTrapezeC"
'
'Revision
'
'Description
'    Definition of Right angle Trapeze
'Notes
'History
'NN      18th May, 2002  Define Concentric Trapeze where two bases are symetrical with three parameters:
'                                   BaseWidth (A), Width (B), and Height (C); and
'                                   9 cardinal points; the origin is the bottom center where is cardinal point 1.
'
'ARV            21st May, 2002  Adjusted location of Cardinal points
'ARV      19th June 2002         Fix for TR 28200
'ARV      21th June 2002        implemented method IsValidParameters
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'ARN       19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const m_progID = "SketchCSSymbol.CTrapezeC"
Const EPSILON = 0.000001

Private Const OUTPUT_1 = "TrapezeC"


Private Const INPUT_1 = "A"
Private Const INPUT_2 = "B"
Private Const INPUT_3 = "C"
Private Const INPUT_4 = "Cardinality"
Private Const INPUT_5 = "Xref"
Private Const INPUT_6 = "Yref"
Private Const E_FAIL = -2147467259

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_EditOccurence = False
    Exit Function
  
ErrorHandler:
    Error.Raise E_FAIL
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
    Error.Raise E_FAIL
End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_progID
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Set m_SymbolDefinition = m_SymbolFactory.CreateEntity(definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition m_SymbolDefinition
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = m_SymbolDefinition
    
    ' Set definition progId and codebase
    m_SymbolDefinition.ProgId = m_progID
    m_SymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    m_SymbolDefinition.Name = m_SymbolDefinition.ProgId
'Arnab        19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'    m_SymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    Exit Function

ErrorHandler:
    Error.Raise E_FAIL
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayofinputs())
    On Error GoTo ErrorHandler
    
    Set m_outputColl = outputcoll
    
    If StrComp(repName, "Symbolic") = 0 Then
        Call Symbolic(arrayofinputs)
    End If
    Exit Sub

ErrorHandler:
    Error.Raise E_FAIL
End Sub

    Private Sub Symbolic(ByRef arrayofinputs())
    On Error GoTo ErrHandler
    Dim lCardinality        As Long
    Dim XCord               As Double
    Dim YCord               As Double
    Dim CCBaseWidth         As Double
    Dim CCWidth             As Double
    Dim CCHeight            As Double
    Dim xRef                As Double
    Dim yRef                As Double
    
    'Check for the validity of inputs
    If IsValidParameters(arrayofinputs) = False Then Exit Sub
    
    CCBaseWidth = arrayofinputs(1)
    CCWidth = arrayofinputs(2)
    CCHeight = arrayofinputs(3)
    lCardinality = arrayofinputs(4)
    xRef = arrayofinputs(5)
    yRef = arrayofinputs(6)
    
    Select Case lCardinality
            
    Case 2: XCord = 0#
            YCord = -CCHeight
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 3: XCord = 0#
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord

    Case 4:  XCord = (CCBaseWidth + CCWidth) / 4
            YCord = -CCHeight / 2
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 5: XCord = -(CCBaseWidth + CCWidth) / 4
            YCord = -CCHeight / 2
            TranslateSymbol arrayofinputs, XCord, YCord

    Case 6: XCord = CCWidth / 2#
            YCord = -CCHeight
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 7: XCord = -CCWidth / 2
            YCord = -CCHeight
            TranslateSymbol arrayofinputs, XCord, YCord

    Case 8: XCord = CCBaseWidth / 2
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 9: XCord = -CCBaseWidth / 2
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
    
    Case 100: XCord = xRef + CCBaseWidth / 2 'User defined
            YCord = yRef
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case Else: XCord = 0
            YCord = -CCHeight / 2
            TranslateSymbol arrayofinputs, XCord, YCord
    
    End Select
    
    Exit Sub
ErrHandler:
    Error.Raise E_FAIL
End Sub
'
' SAMPLE CODE TO CREATE A Trapeze right angle SYMBOL
Private Function Graphics(ByRef arrayofinputs()) As LineString3d
   
    On Error GoTo ErrHandler
    Dim oGeomFactory        As New GeometryFactory
    Dim oLineString         As LineString3d
    Dim dBaseWidth          As Double
    Dim dWidth              As Double
    Dim dHeight             As Double
    Dim dPoints(0 To 14)    As Double
   
    
    dBaseWidth = arrayofinputs(1)
    dWidth = arrayofinputs(2)
    dHeight = arrayofinputs(3)
    
    dPoints(0) = -dBaseWidth / 2
    dPoints(1) = 0#
    dPoints(2) = 0#
    
    dPoints(3) = dBaseWidth / 2
    dPoints(4) = 0#
    dPoints(5) = 0#
    
    dPoints(6) = dWidth / 2
    dPoints(7) = dHeight
    dPoints(8) = 0#
    
    dPoints(9) = -dWidth / 2
    dPoints(10) = dHeight
    dPoints(11) = 0#
    
    'close it
    dPoints(12) = dPoints(0)
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(m_outputColl.ResourceManager, 5, dPoints)
    Set Graphics = oLineString

    Set oLineString = Nothing
    Set oGeomFactory = Nothing

    Exit Function
ErrHandler:
    Error.Raise E_FAIL
End Function


Private Sub TranslateSymbol(ByRef arrayofinputs(), ByVal XTrans As Double, ByVal YTrans As Double)
On Error GoTo ErrHandler
    Dim oTrans4x4           As IJDT4x4
    Dim oVector             As IJDVector

    Dim oLineString         As LineString3d
    
    Set oTrans4x4 = New DT4x4
    Set oVector = New DVector
    
    oVector.Set XTrans, YTrans, 0#
    oTrans4x4.LoadIdentity
    oTrans4x4.Translate oVector
    
    Set oLineString = Graphics(arrayofinputs)
    
    ' if symbol is not there just exit
    If oLineString Is Nothing Then Exit Sub
    
    oLineString.Transform oTrans4x4

    m_outputColl.AddOutput OUTPUT_1, oLineString
    
    Set oLineString = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    Exit Sub
ErrHandler:
    Error.Raise E_FAIL
End Sub


Sub SymbolDefinition(ByRef pSymbolDefinition As Object)
  
    On Error GoTo ErrorHandler
    
    ' Create a new input by new operator
    Dim oInput(1 To 6) As IMSSymbolEntities.IJDInput
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Dim iNumInput As Integer
    Dim iNumOutput As Integer
    Dim iNumRep As Integer
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
  
    ' Create a defaultValue
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    PC.Type = igValue
    
    For iNumInput = 1 To 6
      Set oInput(iNumInput) = New IMSSymbolEntities.DInput
    Next
      
    oInput(1).Name = INPUT_1
    oInput(1).Description = "BaseWidthA"
    oInput(1).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1#
    oInput(1).DefaultParameterValue = PC
    
    oInput(2).Name = INPUT_2
    oInput(2).Description = "WidthB"
    oInput(2).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1#
    oInput(2).DefaultParameterValue = PC
    
    oInput(3).Name = INPUT_3
    oInput(3).Description = "HeightC"
    oInput(3).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1#
    oInput(3).DefaultParameterValue = PC
   
    oInput(4).Name = INPUT_4
    oInput(4).Description = "Cardinality"
    oInput(4).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1
    oInput(4).DefaultParameterValue = PC
    
    oInput(5).Name = INPUT_5
    oInput(5).Description = "Xref"
    oInput(5).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    oInput(5).DefaultParameterValue = PC
    
    oInput(6).Name = INPUT_6
    oInput(6).Description = "Yref"
    oInput(6).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    oInput(6).DefaultParameterValue = PC
    
    ' Set the input to the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    For iNumInput = 1 To 6
      oInputs.SetInput oInput(iNumInput), iNumInput
    Next

    ' Create the output
    Dim oOutput(1 To 1) As IMSSymbolEntities.IJDOutput
    
    For iNumOutput = 1 To 1
      Set oOutput(iNumOutput) = New IMSSymbolEntities.DOutput
    Next
    
    oOutput(1).Name = OUTPUT_1
    oOutput(1).Description = "TrapezeC"
    oOutput(1).Properties = igDESCRIPTION_INIT

  ' Define the representation "Symbolic"
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation

    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = pSymbolDefinition
    
    Set oOutputs = oRep
    
    oRep.Name = "Symbolic"
    oRep.Description = "Description for Rep_1"
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
    Dim dBaseWidth          As Double
    Dim dWidth              As Double
    Dim dHeight             As Double
    Dim lCardinality As Long
    Dim oErrors As IJEditErrors
    Set oErrors = New JServerErrors
    
    dBaseWidth = arrayofinputs(1)
    dWidth = arrayofinputs(2)
    dHeight = arrayofinputs(3)
    lCardinality = arrayofinputs(4)
    IsValidParameters = True
  ' Fix For TR 28200
    'If Symbolattribute value is negative then give message to the user
    If dBaseWidth < EPSILON Or dWidth < EPSILON _
    Or dHeight < EPSILON Or lCardinality < 1 Then
         IsValidParameters = False
         oErrors.Add E_FAIL, "SketchCSSymbol.CTrapezeC", "CrossSection Dimensions are not valid", "TrapezeC"
         GoTo ErrorHandler
    End If

  Exit Function
  
ErrorHandler:
    Error.Raise E_FAIL
End Function


