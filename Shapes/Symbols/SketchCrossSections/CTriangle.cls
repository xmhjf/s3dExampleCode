VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    "SketchCSSymbol.CTriangle"
'
'Revision
'
'Description
'    Definition of Triangle
'Notes
'
'History
'ARN       13th Feb 02           Creation
'ARN        10th April 02         Added CacheOption in
'                                    IJDUserSymbolServices_InstanciateDefinition
'                                    for solving Caching Problem
'ARV        6th May 02            Implemented 4 & 5 Cardinal Points
'ARV        13th May 02           Error handling for invalid input
'ARV       19th June 2002         Fix for TR 28200
'ARV      21th June 2002        implemented method IsValidParameters
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'ARN       19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const m_progID = "SketchCSSymbol.CTriangle"
Const EPSILON = 0.000001
Private Const E_FAIL = -2147467259
Private Const INPUT_1 = "A"
Private Const INPUT_2 = "B"
Private Const INPUT_3 = "C"
Private Const INPUT_4 = "Cardinality"
Private Const INPUT_5 = "Xref"
Private Const INPUT_6 = "Yref"

Private Const OUTPUT_1 = "Triangle"

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
    Dim dSide1              As Double
    Dim dSide2              As Double
    Dim dSide3              As Double
    Dim CosC                As Double
    Dim SinC                As Double
    Dim xRef                As Double
    Dim yRef                As Double
    
    'Check for the validity of inputs
    If IsValidParameters(arrayofinputs) = False Then GoTo ErrHandler
    
    dSide1 = arrayofinputs(1)
    dSide2 = arrayofinputs(2)
    dSide3 = arrayofinputs(3)
    xRef = arrayofinputs(5)
    yRef = arrayofinputs(6)
            
    lCardinality = arrayofinputs(4)
              
    Select Case lCardinality
    Case 2: CosC = (dSide1 * dSide1 + dSide3 * dSide3 - dSide2 * dSide2) / (2 * dSide1 * dSide3)
            SinC = Sqr(1 - CosC * CosC)
            XCord = -dSide3 * CosC
            YCord = -dSide3 * SinC
            TranslateSymbol arrayofinputs, XCord, YCord

    Case 3: XCord = -dSide1 / 2#
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 4: XCord = 0#
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 5: XCord = -dSide1
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 100: XCord = xRef 'User defined
            YCord = yRef
            TranslateSymbol arrayofinputs, XCord, YCord
    
    Case Else: CosC = (dSide1 * dSide1 + dSide3 * dSide3 - dSide2 * dSide2) / (2 * dSide1 * dSide3)
            SinC = Sqr(1 - CosC * CosC)
            XCord = -(dSide3 * CosC + dSide1) / 3#
            YCord = -dSide3 * SinC / 3#
            TranslateSymbol arrayofinputs, XCord, YCord

    End Select
    Exit Sub
ErrHandler:
    Err.Raise E_FAIL
End Sub
'
' SAMPLE CODE TO CREATE A Triangle SYMBOL
Private Function Graphics(ByRef arrayofinputs()) As LineString3d
   On Error GoTo ErrHandler
    Dim oGeomFactory            As New GeometryFactory
    Dim oTriangle               As LineString3d
    Dim dSide1                  As Double
    Dim dSide2                  As Double
    Dim dSide3                  As Double
    Dim dPoints(1 To 12)        As Double
    Dim NoPoints                As Long
    Dim CosC                    As Double
    Dim SinC                    As Double
    Dim lCardinality            As Long
    
    NoPoints = 4
    dSide1 = arrayofinputs(1)
    dSide2 = arrayofinputs(2)
    dSide3 = arrayofinputs(3)

    
    CosC = (dSide1 * dSide1 + dSide3 * dSide3 - dSide2 * dSide2) / (2 * dSide1 * dSide3)
    SinC = Sqr(1 - CosC * CosC)
    'Taking the first vertex as origin i.e vertex B. Constructing Triangle BAC
    dPoints(1) = 0#
    dPoints(2) = 0#
    dPoints(3) = 0#
    dPoints(4) = dSide3 * CosC
    dPoints(5) = dSide3 * SinC
    dPoints(6) = 0#
    dPoints(7) = dSide1
    dPoints(8) = 0#
    dPoints(9) = 0#
    dPoints(10) = 0#
    dPoints(11) = 0#
    dPoints(12) = 0#
   
    Set oTriangle = oGeomFactory.LineStrings3d.CreateByPoints(m_outputColl.ResourceManager, NoPoints, dPoints)
    Set Graphics = oTriangle
    
    Set oTriangle = Nothing
    Set oGeomFactory = Nothing


    Exit Function
ErrHandler:
  Err.Raise E_FAIL
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
  Err.Raise E_FAIL
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
    oInput(1).Description = "SideA"
    oInput(1).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1#
    oInput(1).DefaultParameterValue = PC
    
    oInput(2).Name = INPUT_2
    oInput(2).Description = "SideB"
    oInput(2).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1#
    oInput(2).DefaultParameterValue = PC
    
    oInput(3).Name = INPUT_3
    oInput(3).Description = "SideC"
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
    Err.Raise E_FAIL
End Sub

Private Function IsValidParameters(ByRef arrayofinputs()) As Boolean
On Error GoTo ErrorHandler
    Dim dRadius As Double
    Dim lCardinality As Long
    Dim dSide1 As Double
    Dim dSide2 As Double
    Dim dSide3 As Double
    Dim oErrors As IJEditErrors
    Set oErrors = New JServerErrors
    
    IsValidParameters = True
    
    dSide1 = arrayofinputs(1)
    dSide2 = arrayofinputs(2)
    dSide3 = arrayofinputs(3)
    lCardinality = arrayofinputs(4)
    
  ' Fix For TR 28200
    'If Symbolattribute value is negative then give message to the user
    If dSide1 < EPSILON Or dSide2 < EPSILON Or dSide3 < EPSILON _
    Or lCardinality < 1 Or dSide1 + dSide2 <= dSide3 Or dSide1 + dSide3 <= dSide2 _
    Or dSide2 + dSide3 <= dSide1 Then
         IsValidParameters = False
         oErrors.Add E_FAIL, "SketchCSSymbol.CTriangle", "CrossSection Dimensions are not valid", "Triangle"
         GoTo ErrorHandler
    End If

  Exit Function
  
ErrorHandler:
    Err.Raise E_FAIL
End Function

