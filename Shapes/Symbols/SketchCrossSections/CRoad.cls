VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    CRoad.cls
'
'Revision
'
'Description
'    Definition of Road
'Notes
'History
'   RI        28th Mar 02           Creation
'Arunav        10th April 02         Added CacheOption in
'                                    IJDUserSymbolServices_InstanciateDefinition
'                                    for solving Caching Problem
'   RI        16th Apr 02          Implemented cardinality
'ARV         6th May 02           Implemented cardianl point 7
'ARV       19th June 2002         Fix for TR 28200
'ARV       21th June 2002        implemented method IsValidParameters
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'ARN       19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Road : "
Private Const sPROGID = "SketchCSSymbol.CRoad"
Private Const Symbolic = "Symbolic"
Const EPSILON = 0.000001
Private Const OUTPUT_1 = "Road"
Private Const E_FAIL = -2147467259

Private Const INPUT_1 = "A"
Private Const INPUT_2 = "E"
Private Const INPUT_3 = "F"
Private Const INPUT_4 = "B"
Private Const INPUT_5 = "D"
Private Const INPUT_6 = "C"
Private Const INPUT_7 = "Cardinality"
Private Const INPUT_8 = "Xref"
Private Const INPUT_9 = "Yref"

Private m_oSymbolFactory         As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_oSymbolDefinition      As IMSSymbolEntities.DSymbolDefinition
Private m_oOutputCollection      As IJDOutputCollection

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
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_GetDefinitionName = sPROGID
    Exit Function
  
ErrorHandler:
    Error.Raise E_FAIL
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolDefinition = m_oSymbolFactory.CreateEntity(definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition m_oSymbolDefinition
   
    ' Set definition ssPROGID and codebase
    m_oSymbolDefinition.ProgId = sPROGID
    m_oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    m_oSymbolDefinition.Name = m_oSymbolDefinition.ProgId
'Arnab        19th Jan 2004         TR-CP·54171  The Sketch Cross Section symbol should have "Cached" property
'    m_oSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolDefinition

    Exit Function
ErrorHandler:
    Error.Raise E_FAIL
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayofinputs())
    On Error GoTo ErrorHandler
    
    Set m_oOutputCollection = outputcoll
    
    'in the case, representation is "Symbolic"
    If StrComp(repName, Symbolic) = 0 Then
        GetSymbolic arrayofinputs
    End If
    Exit Sub
ErrorHandler:
    Error.Raise E_FAIL
End Sub

'***************************************************************************
'
'Name: GetSymbolic(ByRef arrayOfInputs())
'
'Parameters : arrayOfInputs() : symbol inputs array
'
'Description : Places the symbol
'              Also translates the symbol based on cardinality
'
'
'****************************************************************************
Private Sub GetSymbolic(ByRef arrayofinputs())
Const METHOD = "GetSymbolic"
On Error GoTo ErrorHandler
    
    Dim lCardinality        As Long
    Dim XCord               As Double
    Dim YCord               As Double
    Dim dA                  As Double
    Dim dE                  As Double
    Dim dF                  As Double
    Dim dB                  As Double
    Dim dD                  As Double
    Dim dC                  As Double
    Dim xRef                As Double
    Dim yRef                As Double
    
    'Check for the validity of inputs
    If IsValidParameters(arrayofinputs) = False Then Exit Sub
    
    'cache all the inputs
    dA = arrayofinputs(1)
    dE = arrayofinputs(2)
    dF = arrayofinputs(3)
    dB = arrayofinputs(4)
    dD = arrayofinputs(5)
    dC = arrayofinputs(6)
    lCardinality = arrayofinputs(7)
    xRef = arrayofinputs(8)
    yRef = arrayofinputs(9)

    'based on the cardinality translate the symbol
    Select Case lCardinality
        Case 2: XCord = -dA * 0.5
            YCord = -dC
            TranslateSymbol arrayofinputs, XCord, YCord
        
        Case 3: XCord = -dA * 0.5
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord

        Case 4: XCord = 0#
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
            
        Case 5: XCord = -dA
            YCord = 0#
            TranslateSymbol arrayofinputs, XCord, YCord
                    
        Case 6: XCord = 0#
            YCord = -dC
            TranslateSymbol arrayofinputs, XCord, YCord
            
        Case 7: XCord = -dA
            YCord = -dE
            TranslateSymbol arrayofinputs, XCord, YCord
        
        Case 100: XCord = xRef 'User defined
            YCord = yRef
            TranslateSymbol arrayofinputs, XCord, YCord
        
        Case Else: XCord = -dA * 0.5
            YCord = -dB
            TranslateSymbol arrayofinputs, XCord, YCord

    End Select
    
    Exit Sub
ErrorHandler:
    Error.Raise E_FAIL
End Sub
'***************************************************************************
'
'Name: Graphics(ByRef arrayOfInputs())
'
'Parameters : arrayOfInputs() : symbol inputs array
'
'Description : constructs the symbol using line string
'
'****************************************************************************
Private Function Graphics(ByRef arrayofinputs()) As LineString3d
Const METHOD = "Graphics"
On Error GoTo ErrorHandler

    Dim oGeomFactory        As New GeometryFactory
    Dim oLineString         As LineString3d
    Dim dParameter2         As Double
    Dim dParameter1         As Double
    Dim dParameter3         As Double
    Dim dParameter4         As Double
    Dim dParameter5         As Double
    Dim dParameter6         As Double
    Dim dPoints(0 To 26)    As Double
   
    'cache the inputs
    dParameter1 = arrayofinputs(1)  '   dA
    dParameter2 = arrayofinputs(2)  '   dE
    dParameter3 = arrayofinputs(3)  '   dF
    dParameter4 = arrayofinputs(4)  '   dB
    dParameter5 = arrayofinputs(5)  '   dD
    dParameter6 = arrayofinputs(6)  '   dC
    
    'construct the points to construct line string
    dPoints(0) = 0#
    dPoints(1) = 0#
    dPoints(2) = 0#
    
    dPoints(3) = dParameter1
    dPoints(4) = 0#
    dPoints(5) = 0#
    
    dPoints(6) = dParameter1
    dPoints(7) = dParameter2
    dPoints(8) = 0#
    
    dPoints(9) = dParameter1 - dParameter3
    dPoints(10) = dParameter2
    dPoints(11) = 0#
    
    dPoints(12) = dParameter1 - dParameter3
    dPoints(13) = dParameter4
    dPoints(14) = 0#
    
    dPoints(15) = dParameter5
    dPoints(16) = dParameter4
    dPoints(17) = 0#
    
    dPoints(18) = dParameter5
    dPoints(19) = dParameter6
    dPoints(20) = 0#
    
    dPoints(21) = 0#
    dPoints(22) = dParameter6
    dPoints(23) = 0#
    
    'close the symbol
    dPoints(24) = dPoints(0)
    dPoints(25) = dPoints(1)
    dPoints(26) = dPoints(2)
    
    'construct the line string based on the above points
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(m_oOutputCollection.ResourceManager, 9, dPoints)
    Set Graphics = oLineString

    Set oLineString = Nothing
    Set oGeomFactory = Nothing

    Exit Function
ErrorHandler:
    Error.Raise E_FAIL
End Function

'***************************************************************************
'
'Name:          TranslateSymbol(ByRef arrayOfInputs(), ByVal XTrans As Double, ByVal YTrans As Double)
'
'Parameters :   arrayOfInputs() :   symbol inputs array
'               XTrans          :   X coordinate of the point where symbol needs to be translated
'               XTrans          :   Y coordinate of the point where symbol needs to be translated
'Description :  Translates the symbol to the given point
'
'****************************************************************************
Private Sub TranslateSymbol(ByRef arrayofinputs(), ByVal XTrans As Double, ByVal YTrans As Double)
Const METHOD = "TranslateSymbol"
On Error GoTo ErrorHandler
    Dim oTrans4x4           As IJDT4x4
    Dim oVector             As IJDVector

    Dim oLineString         As LineString3d
    Set oTrans4x4 = New DT4x4
    Set oVector = New DVector
    
    oVector.Set XTrans, YTrans, 0#
    'construct identity matrix
    oTrans4x4.LoadIdentity
    
    'load the matrix with translation data
    oTrans4x4.Translate oVector
    
    'get the line string
    Set oLineString = Graphics(arrayofinputs)
    ' if symbol is not there just exit
    If oLineString Is Nothing Then Exit Sub
    
    'transform the line string
    oLineString.Transform oTrans4x4
    
    'add the output to the collection
    m_oOutputCollection.AddOutput OUTPUT_1, oLineString
    
    Set oLineString = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    Exit Sub
ErrorHandler:
    Error.Raise E_FAIL
End Sub

'***************************************************************************
'
'Name:          SymbolDefinition(ByRef pSymbolDefinition As Object)
'
'Parameters :   pSymbolDefinition :   symbol object
'
'Description :  defines the symbol data like inputs, representation and
'               representation evaluation required to construct it
'
'****************************************************************************

Sub SymbolDefinition(ByRef pSymbolDefinition As Object)
Const METHOD = "SymbolDefinition"
On Error GoTo ErrorHandler
    
    Dim oInput(1 To 9)          As IMSSymbolEntities.IJDInput
    Dim oParameterContent       As IMSSymbolEntities.IJDParameterContent
    Dim iNumInput               As Integer
    Dim iNumOutput              As Integer
    Dim iNumRep                 As Integer
    Dim oOutputs                As IMSSymbolEntities.IJDOutputs
  
    ' Create a defaultValue
    Set oParameterContent = New IMSSymbolEntities.DParameterContent 'not persistent oParameterContent
    oParameterContent.Type = igValue
    
    For iNumInput = 1 To 9
      Set oInput(iNumInput) = New IMSSymbolEntities.DInput
    Next
      
    'Set the symbol input data
    oInput(1).Name = INPUT_1
    oInput(1).Description = "A"
    oInput(1).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 4#
    oInput(1).DefaultParameterValue = oParameterContent

    oInput(2).Name = INPUT_2
    oInput(2).Description = "E"
    oInput(2).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 0.4
    oInput(2).DefaultParameterValue = oParameterContent

    oInput(3).Name = INPUT_3
    oInput(3).Description = "F"
    oInput(3).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 0.2
    oInput(3).DefaultParameterValue = oParameterContent
    
    oInput(4).Name = INPUT_4
    oInput(4).Description = "B"
    oInput(4).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 0.15
    oInput(4).DefaultParameterValue = oParameterContent

    oInput(5).Name = INPUT_5
    oInput(5).Description = "D"
    oInput(5).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 0.27
    oInput(5).DefaultParameterValue = oParameterContent
    
    oInput(6).Name = INPUT_6
    oInput(6).Description = "C"
    oInput(6).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 0.45
    oInput(6).DefaultParameterValue = oParameterContent

    oInput(7).Name = INPUT_7
    oInput(7).Description = "Cardinality"
    oInput(7).Properties = igINPUT_IS_A_PARAMETER
    oParameterContent.UomValue = 1
    oInput(7).DefaultParameterValue = oParameterContent

    oInput(8).Name = INPUT_8
    oInput(8).Description = "Xref"
    oInput(8).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    oParameterContent.UomValue = 0
    oInput(8).DefaultParameterValue = oParameterContent
    
    oInput(9).Name = INPUT_9
    oInput(9).Description = "Yref"
    oInput(9).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    oParameterContent.UomValue = 0
    oInput(9).DefaultParameterValue = oParameterContent
    
    ' Set the input to the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    For iNumInput = 1 To 9
      oInputs.SetInput oInput(iNumInput), iNumInput
    Next

    ' Create the output
    Dim oOutput(1 To 1) As IMSSymbolEntities.IJDOutput
    
    For iNumOutput = 1 To 1
      Set oOutput(iNumOutput) = New IMSSymbolEntities.DOutput
    Next
    
    oOutput(1).Name = OUTPUT_1
    oOutput(1).Description = "Road"
    oOutput(1).Properties = igDESCRIPTION_INIT

  ' Define the representation "Symbolic"
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation

    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = pSymbolDefinition
    
    Set oOutputs = oRep
    
    oRep.Name = Symbolic
    oRep.Description = "Description for Rep_1"
    oRep.Properties = igREPRESENTATION_ISVBFUNCTION
    oRep.RepresentationId = 1

    oOutputs.SetOutput oOutput(1)
    oReps.SetRepresentation oRep

    'Set the symbol representation evaluation
    Dim oVbFuncRep(1 To 1) As IJDRepresentationEvaluation
    
    For iNumRep = 1 To 1
      Set oVbFuncRep(iNumRep) = New DRepresentationEvaluation
      oVbFuncRep(iNumRep).Properties = igREPRESENTATION_HIDDEN
      oVbFuncRep(iNumRep).Type = igREPRESENTATION_VBFUNCTION
      oVbFuncRep(iNumRep).ProgId = sPROGID
    Next iNumRep
    
    Dim oScripts As IMSSymbolEntities.IJDRepresentationEvaluations
    Set oScripts = pSymbolDefinition
    
    oVbFuncRep(1).Name = Symbolic
    oVbFuncRep(1).Description = "Script for the Symbolic representation"
    oScripts.AddRepresentationEvaluation oVbFuncRep(1)

    Exit Sub
ErrorHandler:
    Error.Raise E_FAIL
End Sub

Private Function IsValidParameters(ByRef arrayofinputs()) As Boolean
On Error GoTo ErrorHandler
    Dim dParameter1         As Double
    Dim dParameter2         As Double
    Dim dParameter3         As Double
    Dim dParameter4         As Double
    Dim dParameter5         As Double
    Dim dParameter6         As Double
    Dim lCardinality As Long
    Dim oErrors As IJEditErrors
    Set oErrors = New JServerErrors
    
    dParameter1 = arrayofinputs(1)
    dParameter2 = arrayofinputs(2)
    dParameter3 = arrayofinputs(3)
    dParameter4 = arrayofinputs(4)
    dParameter5 = arrayofinputs(5)
    dParameter6 = arrayofinputs(6)
    
    lCardinality = arrayofinputs(7)
    
    IsValidParameters = True
  ' Fix For TR 28200
    'If Symbolattribute value is negative then give message to the user
    If dParameter1 < EPSILON Or dParameter2 < EPSILON _
    Or dParameter3 < EPSILON Or dParameter4 < EPSILON _
    Or dParameter5 < EPSILON Or dParameter6 < EPSILON _
    Or lCardinality < 1 Then
         IsValidParameters = False
         oErrors.Add E_FAIL, "SketchCSSymbol.CRoad", "CrossSection Dimensions are not valid", "Road"
         GoTo ErrorHandler
    End If

  Exit Function
  
ErrorHandler:
    Error.Raise E_FAIL
End Function


