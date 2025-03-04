VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEllipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    "SketchCSSymbol.CEllipse"
'
'Revision
'
'Description
'    Definition of Ellipse
'Notes
'History
''ARV        11th Feb 02           Creation
'ARV        10th April 02         Added CacheOption in
'                                    IJDUserSymbolServices_InstanciateDefinition
'                                    for solving Caching Problem
'ARV         6th May 02           Corrected Cardinality
'ARV       19th June 2002        Fix for TR 28200
'ARV    21th June 2002        implemented method IsValidParameters
'NN          4 Aug 2002          Assign proper inputs, define axis ratio, and construct correct ellipse and define Cardinal Points.
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'ARN        19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const m_progID = "SketchCSSymbol.CEllipse"
Const EPSILON = 0.000001
Private Const E_FAIL = -2147467259
Private Const OUTPUT_1 = "ELLIPSE"

Private Const INPUT_1 = "A"
Private Const INPUT_2 = "B"
Private Const INPUT_3 = "Cardinality"
Private Const INPUT_4 = "Xref"
Private Const INPUT_5 = "Yref"


Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
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
    Err.Raise E_FAIL

End Sub

Private Sub Symbolic(ByRef arrayofinputs())
On Error GoTo ErrHandler

    Dim lCardinality        As Long
    Dim dMajorRadius        As Double
    Dim dMinorRadius        As Double
    Dim XCord               As Double
    Dim YCord               As Double
    Dim xRef                As Double
    Dim yRef                As Double
    
    'Check for the validity of inputs
    If IsValidParameters(arrayofinputs) = False Then Exit Sub
    
    dMajorRadius = arrayofinputs(1)
    dMinorRadius = arrayofinputs(2)
    lCardinality = arrayofinputs(3)
    xRef = arrayofinputs(4)
    yRef = arrayofinputs(5)
 
    Select Case lCardinality
    Case 2: XCord = 0#
            YCord = -dMinorRadius
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 3: XCord = 0#
            YCord = dMinorRadius
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 4: XCord = dMajorRadius
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
    Case 5: XCord = -dMajorRadius
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
    
    Case 100: XCord = xRef - dMajorRadius 'User defined
        YCord = yRef
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
' SAMPLE CODE TO CREATE A Ellipse SYMBOL
Private Function Graphics(ByRef arrayofinputs()) As Ellipse3d
  
    On Error GoTo ErrHandler
    Dim oGeomFactory        As New GeometryFactory
    Dim oEllipse            As Ellipse3d
    Dim dMajorRadius        As Double
    Dim dMinorRadius        As Double
    Dim dMMRatio            As Double

    
    dMajorRadius = arrayofinputs(1)
    dMinorRadius = arrayofinputs(2)
    dMMRatio = dMinorRadius / dMajorRadius
    
    Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, 0, 0, 0, _
                    0, 0, 1, dMajorRadius, 0, 0, dMMRatio)
                    
    Set Graphics = oEllipse

    Set oEllipse = Nothing
    Set oGeomFactory = Nothing

    Exit Function
ErrHandler:
    Error.Raise E_FAIL
End Function
Private Sub TranslateSymbol(ByRef arrayofinputs(), XTrans As Double, YTrans As Double)
On Error GoTo ErrHandler
    Dim oTrans4x4           As IJDT4x4
    Dim oVector             As IJDVector

    Dim oEllipse            As Ellipse3d
    
    Set oTrans4x4 = New DT4x4
    Set oVector = New DVector
    
    oVector.Set XTrans, YTrans, 0#
    oTrans4x4.LoadIdentity
    oTrans4x4.Translate oVector
    
    Set oEllipse = Graphics(arrayofinputs)
    ' if symbol is not there just exit
    If oEllipse Is Nothing Then Exit Sub
    
    oEllipse.Transform oTrans4x4

    m_outputColl.AddOutput OUTPUT_1, oEllipse
    
    Set oEllipse = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    
    Exit Sub
ErrHandler:
    Error.Raise E_FAIL
End Sub

Sub SymbolDefinition(ByRef pSymbolDefinition As Object)
  
    On Error GoTo ErrorHandler
    
    ' Create a new input by new operator
    Dim Coord(1 To 5) As IMSSymbolEntities.IJDInput
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Dim iNumInput As Integer
    Dim iNumOutput As Integer
    Dim iNumRep As Integer
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
  
    ' Create a defaultValue
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    PC.Type = igValue
    
    For iNumInput = 1 To 5
      Set Coord(iNumInput) = New IMSSymbolEntities.DInput
    Next

    Coord(1).Name = INPUT_1
    Coord(1).Description = "Major Radius"
    Coord(1).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1
    Coord(1).DefaultParameterValue = PC

    Coord(2).Name = INPUT_2
    Coord(2).Description = "Minor Radius"
    Coord(2).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1
    Coord(2).DefaultParameterValue = PC

    Coord(3).Name = INPUT_3
    Coord(3).Description = "Cardinality"
    Coord(3).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 1
    Coord(3).DefaultParameterValue = PC
    
    Coord(4).Name = INPUT_4
    Coord(4).Description = "Xref"
    Coord(4).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Coord(4).DefaultParameterValue = PC
    
    Coord(5).Name = INPUT_5
    Coord(5).Description = "Yref"
    Coord(5).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Coord(5).DefaultParameterValue = PC
     
    ' Set the input to the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    For iNumInput = 1 To 5
      oInputs.SetInput Coord(iNumInput), iNumInput
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
    
    Dim dMajorRadius        As Double
    Dim dMinorRadius        As Double
    Dim lCardinality As Long
    Dim oErrors As IJEditErrors
    Set oErrors = New JServerErrors
    
    dMajorRadius = arrayofinputs(2)
    dMinorRadius = arrayofinputs(1)
    lCardinality = arrayofinputs(3)
    
    IsValidParameters = True
  ' Fix For TR 28200
    'If Symbolattribute value is negative then give message to the user
    If dMajorRadius < EPSILON Or dMinorRadius < EPSILON Or lCardinality < 1 Then
         IsValidParameters = False
         oErrors.Add E_FAIL, "SketchCSSymbol.CEllipse", "CrossSection Dimensions are not valid", "Ellipse"
         GoTo ErrorHandler
    End If

  Exit Function
  
ErrorHandler:
    Error.Raise E_FAIL
End Function

