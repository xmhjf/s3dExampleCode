VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDefFndPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'File
'    UserDefFndPort.cls
'
'Author
'    RMP
'
'Description
'    USS Implementation for a foundation port with a "Non Standard" bolthole
'    pattern, the bolt hole coordinates are specified by the user.
'   Change History:
'   dd.mmm.yyyy        who                           change description
'   -----------       -----                          ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Implements IJDUserSymbolServices
Dim m_outputColl As IJDOutputCollection
Private Const E_FAIL = -2147467259


Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "SP3DFoundationPorts.UserDefFndPort"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(ppSymbolDefDisp As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    ppSymbolDefDisp.IJDInputs.RemoveAllInput
    ppSymbolDefDisp.IJDRepresentations.RemoveAllRepresentation
    ppSymbolDefDisp.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    FeedPortDefinition ppSymbolDefDisp
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
    Dim oSymbolFactory As IMSSymbolEntities.IJDSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    
    Set oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, pResourceMgr)
    Set oSymbolFactory = Nothing
    
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "SP3DFoundationPorts.UserDefFndPort"
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    Set m_outputColl = pOutputColl
    If StrComp(pRepName, "Physical") = 0 Then
        Symbolic arrayOfInputs
    End If
End Sub

Private Sub Symbolic(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Dim lNumberOfBoltholes As Long
    Dim strHoleCoordinates As String
    Dim lsPoints() As Double
    Dim oBoltholesCurve As IJCurve, oPlane As Object
    Dim oLineString As IJLineString
    Dim lArrayUpperDim As Long
    Dim oPoint3D As Point3d
        
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    lNumberOfBoltholes = arrayOfInputs(1)
    strHoleCoordinates = arrayOfInputs(2)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    
    ' Check for error conditions here and log errors
        'oErrors.Add E_FAIL, "SP3DFoundationPorts.FlangedNozzle", "Nozzle matching dimensional data does not exist. Resetting the values.", "flangeDia <= 0 Or pipeDia <= 0"
        ' GoTo ErrorHandler
    ' End If

    If lNumberOfBoltholes > 1 Then
        lArrayUpperDim = (lNumberOfBoltholes * 3) - 1
        ReDim lsPoints(0 To lArrayUpperDim)
        
        FillPointsArray strHoleCoordinates, lNumberOfBoltholes, lsPoints
        Set oLineString = geomFactory.LineStrings3d.CreateByPoints( _
                                m_outputColl.ResourceManager, lNumberOfBoltholes, lsPoints)
        Set oBoltholesCurve = oLineString
        
        ' If the curve is not a linear thing (degenerated to a straight line), close the curve if
        ' it is not closed and attempt to create a plane bounded by it.
        ' If the curve has degenerated into a line, we do not create a bounded plane at all.
        If oBoltholesCurve.Scope <> CURVE_SCOPE_COLINEAR Then
            ' Check whether the linestring is closed, plane boundaries have to be closed curves.
            ' If it is not closed, close it. Math internally handles adding the extra point.
            If oLineString.IsClosed = False Then
                oLineString.IsClosed = True
            End If
        
            Set oPlane = geomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, oBoltholesCurve)
            m_outputColl.AddOutput "BoundedPlane", oPlane
        End If
        
        m_outputColl.AddOutput "BoltHolesCurve", oBoltholesCurve
    Else
        ' Create a point at the origin when there are one or no holes
        Set oPoint3D = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, 0, 0, 0)
            
        If lNumberOfBoltholes = 1 Then
            
            lArrayUpperDim = (lNumberOfBoltholes * 3) - 1
            ReDim lsPoints(0 To lArrayUpperDim)
            
            FillPointsArray strHoleCoordinates, lNumberOfBoltholes, lsPoints
        
            ' Set the position at the given location
            oPoint3D.SetPoint lsPoints(0), lsPoints(1), lsPoints(2)
        End If
        
        m_outputColl.AddOutput "BoltHolesCurve", oPoint3D
    End If

    Set oPoint3D = Nothing
    Set oPlane = Nothing
    Set oBoltholesCurve = Nothing
    
    Set geomFactory = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'
'   This symbol has two inputs:
'       1. Number of Boltholes
'       2. String with all hole coordinates serialized into it.
'   The number of outputs is variable.
'       * If the boltholes form a closed curve, we'll have two outputs: the curve
'           itself and a plane bounded by this curve
'       * If the boltholes do not form a closed curve, there will only be one output:
'           the curve itself.
'
Private Sub FeedPortDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
  
    On Error GoTo ErrorHandler
    
    ' Create a new input by new operator
    Dim Inputs(1 To 2) As IMSSymbolEntities.IJDInput
    
    ' Create a defaultValue
    Dim PC1 As IJDParameterContent, PC2 As IJDParameterContent
    Set PC1 = New DParameterContent 'not persistent PC
    Set PC2 = New DParameterContent 'not persistent PC
    
    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "NumberOfHoles"
    Inputs(1).Description = "Number of Bolt Holes"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igValue
    PC1.UomValue = 5
    Inputs(1).DefaultParameterValue = PC1
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "HoleCoordinates"
    Inputs(2).Description = "Coordinates of the holes"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC2.Type = igString
    PC2.String = "0.0|0.0|0.5|0.0|0.5|0.5|0.0|0.5|0.0|0.0"
    Inputs(2).DefaultParameterValue = PC2

    'Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    Dim index As Integer
    For index = 1 To 2
      InputsIf.SetInput Inputs(index), index
    Next
    For index = 1 To 2
        Set Inputs(index) = Nothing
    Next
    
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.IJDRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
    
    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of Port"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    ' The number of outputs is variable
    rep1.Properties = igCOLLECTION_VARIABLE
    rep1.RepresentationId = SimplePhysical
    
    'Set the 'Physical' representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
    
    Set rep1 = Nothing
    Set RepsIf = Nothing
    
    'Define the evaluation for the Physical representation
    Dim SymbolicRepEval As IJDRepresentationEvaluation
    Set SymbolicRepEval = New DRepresentationEvaluation
    SymbolicRepEval.Name = "Physical"
    SymbolicRepEval.Description = "Physical representation of Foundation Port"
    SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
    SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
    SymbolicRepEval.ProgId = "SP3DFoundationPorts.UserDefFndPort"
    
    'Set the evaluations on symbol definition
    Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition
    
    RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval
    
    Set RepsEvalsIf = Nothing
    Set SymbolicRepEval = Nothing
    
    Exit Sub
ErrorHandler:
    Error.Raise E_FAIL
End Sub


'
'Notes:
'   "0.0|0.0|0.5|0.0|0.5|0.5|0.0|0.5|0.0|0.0"
'
Private Sub FillPointsArray(strSerializedCoordinates As String, lNumHoles As Long, _
                                lsPoints() As Double)
    
    Dim strTemp As String, strRem As String
    Dim li As Long, lPos As Long, lStrLength As Long
    
    ' allocate a copy of the given string
    strTemp = strSerializedCoordinates
    
    For li = 0 To lNumHoles - 1
        ' The Z Coordinate, negative of the passed in X coordinate
        lPos = InStr(1, strTemp, "|", vbTextCompare)
        lsPoints(li * 3 + 2) = CDbl(Left(strTemp, lPos - 1))
        lStrLength = Len(strTemp)
        strRem = Right(strTemp, lStrLength - lPos)
        strTemp = strRem
        strRem = vbNullString
        
        ' The Y Coordinate
        If li = lNumHoles - 1 Then
            lsPoints(li * 3 + 1) = CDbl(strTemp)
        Else
            lPos = InStr(1, strTemp, "|", vbTextCompare)
            lsPoints(li * 3 + 1) = CDbl(Left(strTemp, lPos - 1))
            lStrLength = Len(strTemp)
            strRem = Right(strTemp, lStrLength - lPos)
            strTemp = strRem
            strRem = vbNullString
        End If
        
        ' The X Coordinate, always Zero
        lsPoints(li * 3) = 0
    Next li
    
End Sub

Private Sub Class_Terminate()
Set m_outputColl = Nothing
End Sub
