VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HexagonalSolidServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    HexagonalSolidServices.cls
'
'Author
'    MU
'
'Description
'    Definition of HexagonalSolidServices Symbol
'History
' 07 April Mule Creation
'   24Jul2003   VS     Do not clear middle tier errors. This may clear out other errors too.
'   15Oct2003   VS      CR36485 - Added outputs for more edges
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const E_FAIL = -2147467259
Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "HexagonalSolid.HexagonalSolidServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    ' Feed HexagonalSolid Definition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "A"
    '          3. "B"
    '          2. "C"
    '          3. "D"

    ' Representations :
    '           Physical
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 6) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'Define inputs for HexagonalSolid
    Dim Index As Integer
    For Index = 1 To 4
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.3
            Inputs(Index).Name = "HexagonalSolidDepth"
            Inputs(Index).Description = "Depth of the Hexagonal solid"
        ElseIf (Index = 2) Then
            PC.UomValue = 0.6
            Inputs(Index).Name = "HexagonalSolidLength1"
        ElseIf (Index = 3) Then
            PC.UomValue = 0.6
            Inputs(Index).Name = "HexagonalSolidLength2"
        ElseIf (Index = 4) Then
            PC.UomValue = 0.3
            Inputs(Index).Name = "HexagonalSolidSide"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 4
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next
    
    'Define the outputs
    Dim oPoint1 As IMSSymbolEntities.DOutput
    Set oPoint1 = New IMSSymbolEntities.DOutput
    
    oPoint1.Name = "Point1"
    oPoint1.Description = "HexagonalSolid"
    oPoint1.Properties = 0
    
    Dim oPoint2 As IMSSymbolEntities.DOutput
    Set oPoint2 = New IMSSymbolEntities.DOutput
    
    oPoint2.Name = "Point2"
    oPoint2.Description = "HexagonalSolid"
    oPoint2.Properties = 0
    
'    Dim oPoint3 As IMSSymbolEntities.DOutput
'    Set oPoint3 = New IMSSymbolEntities.DOutput
'    oPoint3.Name = "Point3"
'    oPoint3.Description = "HexagonalSolid"
'    oPoint3.Properties = 0
'
'    Dim oPoint4 As IMSSymbolEntities.DOutput
'    Set oPoint4 = New IMSSymbolEntities.DOutput
'    oPoint4.Name = "Point4"
'    oPoint4.Description = "HexagonalSolid"
'    oPoint4.Properties = 0
'
'    Dim oPoint5 As IMSSymbolEntities.DOutput
'    Set oPoint5 = New IMSSymbolEntities.DOutput
'    oPoint5.Name = "Point5"
'    oPoint5.Description = "HexagonalSolid"
'    oPoint5.Properties = 0
'
'    Dim oPoint6 As IMSSymbolEntities.DOutput
'    Set oPoint6 = New IMSSymbolEntities.DOutput
'    oPoint6.Name = "Point6"
'    oPoint6.Description = "HexagonalSolid"
'    oPoint6.Properties = 0
'
'    Dim oPoint7 As IMSSymbolEntities.DOutput
'    Set oPoint7 = New IMSSymbolEntities.DOutput
'    oPoint7.Name = "Point7"
'    oPoint7.Description = "HexagonalSolid"
'    oPoint7.Properties = 0
'
'    Dim oPoint8 As IMSSymbolEntities.DOutput
'    Set oPoint8 = New IMSSymbolEntities.DOutput
'    oPoint8.Name = "Point8"
'    oPoint8.Description = "HexagonalSolid"
'    oPoint8.Properties = 0
'
    Dim oTopSurface As IMSSymbolEntities.DOutput
    Set oTopSurface = New IMSSymbolEntities.DOutput
    
    oTopSurface.Name = "TopSurface"
    oTopSurface.Description = "TopSurface of HexagonalSolid"
    oTopSurface.Properties = 0
    
    Dim oBottomSurface As IMSSymbolEntities.DOutput
    Set oBottomSurface = New IMSSymbolEntities.DOutput
    
    oBottomSurface.Name = "BottomSurface"
    oBottomSurface.Description = "BottomSurface of HexagonalSolid"
    oBottomSurface.Properties = 0
    
    Dim oSurface1 As IMSSymbolEntities.DOutput
    Set oSurface1 = New IMSSymbolEntities.DOutput
    
    oSurface1.Name = "Surface1"
    oSurface1.Description = "Surface1 of HexagonalSolid"
    oSurface1.Properties = 0
    
    Dim oSurface2 As IMSSymbolEntities.DOutput
    Set oSurface2 = New IMSSymbolEntities.DOutput
    
    oSurface2.Name = "Surface2"
    oSurface2.Description = "Surface2 of HexagonalSolid"
    oSurface2.Properties = 0
    
    Dim oSurface3 As IMSSymbolEntities.DOutput
    Set oSurface3 = New IMSSymbolEntities.DOutput
    
    oSurface3.Name = "Surface3"
    oSurface3.Description = "Surface3 of HexagonalSolid"
    oSurface3.Properties = 0
    
    Dim oSurface4 As IMSSymbolEntities.DOutput
    Set oSurface4 = New IMSSymbolEntities.DOutput
    
    oSurface4.Name = "Surface4"
    oSurface4.Description = "Surface4 of HexagonalSolid"
    oSurface4.Properties = 0
    
    Dim oSurface5 As IMSSymbolEntities.DOutput
    Set oSurface5 = New IMSSymbolEntities.DOutput
    
    oSurface5.Name = "Surface5"
    oSurface5.Description = "Surface5 of HexagonalSolid"
    oSurface5.Properties = 0
    
    Dim oSurface6 As IMSSymbolEntities.DOutput
    Set oSurface6 = New IMSSymbolEntities.DOutput
    
    oSurface6.Name = "Surface6"
    oSurface6.Description = "Surface6 of HexagonalSolid"
    oSurface6.Properties = 0
    
    Dim E1 As IMSSymbolEntities.IJDOutput
    Set E1 = New IMSSymbolEntities.DOutput
    E1.Name = "Edge1"
    E1.Description = "Edge1 of HexagonalSolid"
    E1.Properties = 0

    Dim E2 As IMSSymbolEntities.IJDOutput
    Set E2 = New IMSSymbolEntities.DOutput
    E2.Name = "Edge2"
    E2.Description = "Edge2 of HexagonalSolid"
    E2.Properties = 0

    Dim E3 As IMSSymbolEntities.IJDOutput
    Set E3 = New IMSSymbolEntities.DOutput
    E3.Name = "Edge3"
    E3.Description = "Edge3 of HexagonalSolid"
    E3.Properties = 0

    Dim E4 As IMSSymbolEntities.IJDOutput
    Set E4 = New IMSSymbolEntities.DOutput
    E4.Name = "Edge4"
    E4.Description = "Edge4 of HexagonalSolid"
    E4.Properties = 0

    Dim E5 As IMSSymbolEntities.IJDOutput
    Set E5 = New IMSSymbolEntities.DOutput
    E5.Name = "Edge5"
    E5.Description = "Edge5 of HexagonalSolid"
    E5.Properties = 0

    Dim E6 As IMSSymbolEntities.IJDOutput
    Set E6 = New IMSSymbolEntities.DOutput
    E6.Name = "Edge6"
    E6.Description = "Edge6 of HexagonalSolid"
    E6.Properties = 0

    Dim E7 As IMSSymbolEntities.IJDOutput
    Set E7 = New IMSSymbolEntities.DOutput
    E7.Name = "Edge7"
    E7.Description = "Edge7 of HexagonalSolid"
    E7.Properties = 0

    Dim E8 As IMSSymbolEntities.IJDOutput
    Set E8 = New IMSSymbolEntities.DOutput
    E8.Name = "Edge8"
    E8.Description = "Edge8 of HexagonalSolid"
    E8.Properties = 0

    Dim E9 As IMSSymbolEntities.IJDOutput
    Set E9 = New IMSSymbolEntities.DOutput
    E9.Name = "Edge9"
    E9.Description = "Edge9 of HexagonalSolid"
    E9.Properties = 0
    
    Dim E10 As IMSSymbolEntities.IJDOutput
    Set E10 = New IMSSymbolEntities.DOutput
    E10.Name = "Edge10"
    E10.Description = "Edge10 of HexagonalSolid"
    E10.Properties = 0
    
    Dim E11 As IMSSymbolEntities.IJDOutput
    Set E11 = New IMSSymbolEntities.DOutput
    E11.Name = "Edge11"
    E11.Description = "Edge11 of HexagonalSolid"
    E11.Properties = 0
    
    Dim E12 As IMSSymbolEntities.IJDOutput
    Set E12 = New IMSSymbolEntities.DOutput
    E12.Name = "Edge12"
    E12.Description = "Edge12 of HexagonalSolid"
    E12.Properties = 0
    
    Dim E13 As IMSSymbolEntities.IJDOutput
    Set E13 = New IMSSymbolEntities.DOutput
    E13.Name = "Edge13"
    E13.Description = "Edge13 of HexagonalSolid"
    E13.Properties = 0
    
    Dim E14 As IMSSymbolEntities.IJDOutput
    Set E14 = New IMSSymbolEntities.DOutput
    E14.Name = "Edge14"
    E14.Description = "Edge14 of HexagonalSolid"
    E14.Properties = 0
    
    Dim E15 As IMSSymbolEntities.IJDOutput
    Set E15 = New IMSSymbolEntities.DOutput
    E15.Name = "Edge15"
    E15.Description = "Edge15 of HexagonalSolid"
    E15.Properties = 0
    
    Dim E16 As IMSSymbolEntities.IJDOutput
    Set E16 = New IMSSymbolEntities.DOutput
    E16.Name = "Edge16"
    E16.Description = "Edge16 of HexagonalSolid"
    E16.Properties = 0
    
    Dim E17 As IMSSymbolEntities.IJDOutput
    Set E17 = New IMSSymbolEntities.DOutput
    E17.Name = "Edge17"
    E17.Description = "Edge17 of HexagonalSolid"
    E17.Properties = 0
    
    Dim E18 As IMSSymbolEntities.IJDOutput
    Set E18 = New IMSSymbolEntities.DOutput
    E18.Name = "Edge18"
    E18.Description = "Edge18 of HexagonalSolid"
    E18.Properties = 0
    
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of HexagonalSolid"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    ' Set the output
    IJDOutputs.SetOutput oTopSurface
    Set oTopSurface = Nothing
    
    IJDOutputs.SetOutput oBottomSurface
    Set oBottomSurface = Nothing
    
    IJDOutputs.SetOutput oSurface1
    Set oSurface1 = Nothing
    
    IJDOutputs.SetOutput oSurface2
    Set oSurface2 = Nothing
    
    IJDOutputs.SetOutput oSurface3
    Set oSurface3 = Nothing
    
    IJDOutputs.SetOutput oSurface4
    Set oSurface4 = Nothing
    
    IJDOutputs.SetOutput oSurface5
    Set oSurface5 = Nothing
    
    IJDOutputs.SetOutput oSurface6
    Set oSurface6 = Nothing
    
    IJDOutputs.SetOutput oPoint1
    Set oPoint1 = Nothing
    
    IJDOutputs.SetOutput oPoint2
    Set oPoint2 = Nothing
'
'    IJDOutputs.SetOutput oPoint3
'    Set oPoint3 = Nothing
'
'    IJDOutputs.SetOutput oPoint4
'    Set oPoint4 = Nothing
'
'    IJDOutputs.SetOutput oPoint5
'    Set oPoint5 = Nothing
'
'    IJDOutputs.SetOutput oPoint6
'    Set oPoint6 = Nothing
'
'    IJDOutputs.SetOutput oPoint7
'    Set oPoint7 = Nothing
'
'    IJDOutputs.SetOutput oPoint8
'    Set oPoint8 = Nothing
    
    IJDOutputs.SetOutput E1
    Set E1 = Nothing

    IJDOutputs.SetOutput E2
    Set E2 = Nothing

    IJDOutputs.SetOutput E3
    Set E3 = Nothing

    IJDOutputs.SetOutput E4
    Set E4 = Nothing

    IJDOutputs.SetOutput E5
    Set E5 = Nothing

    IJDOutputs.SetOutput E6
    Set E6 = Nothing

    IJDOutputs.SetOutput E7
    Set E7 = Nothing

    IJDOutputs.SetOutput E8
    Set E8 = Nothing

    IJDOutputs.SetOutput E9
    Set E9 = Nothing

    IJDOutputs.SetOutput E10
    Set E10 = Nothing

    IJDOutputs.SetOutput E11
    Set E11 = Nothing

    IJDOutputs.SetOutput E12
    Set E12 = Nothing

    IJDOutputs.SetOutput E13
    Set E13 = Nothing

    IJDOutputs.SetOutput E14
    Set E14 = Nothing

    IJDOutputs.SetOutput E15
    Set E15 = Nothing

    IJDOutputs.SetOutput E16
    Set E16 = Nothing

    IJDOutputs.SetOutput E17
    Set E17 = Nothing

    IJDOutputs.SetOutput E18
    Set E18 = Nothing
    
    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    Set rep1 = Nothing
    Set RepsIf = Nothing
    Set IJDOutputs = Nothing
    
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As IJDRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "Physical representation of HexagonalSolid"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "HexagonalSolid.HexagonalSolidServices"
    
    'Set the evaluations for the Symbolic and Physical representations on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    Set PhysicalRepEval = Nothing
    Set RepEvalsIf = Nothing
    
'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    Exit Sub

ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "HexagonalSolid.HexagonalSolidServices"
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
    
    Exit Function

ErrorHandler:
     Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal RepName As String, ByVal OutputColl As Object, arrayOfInputs() As Variant)
    On Error GoTo ErrorHandler
    Set m_outputColl = OutputColl
    If StrComp(RepName, "Physical") = 0 Then
        Physical arrayOfInputs
    End If
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HexagonalSolid
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Dim a  As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    
    a = arrayOfInputs(2)
    b = arrayOfInputs(3)
    c = arrayOfInputs(4)
    d = arrayOfInputs(5)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If a <= 0 Or b <= 0 Or c <= 0 Or d <= 0 Then
        oErrors.Add E_FAIL, "HexagonalSolid.HexagonalSolidServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    ElseIf (d >= c) Then
        'For Regular Hexagonal Solid b=c=d/2
        oErrors.Add E_FAIL, "HexagonalSolid.HexagonalSolidServices", "D should not be greater than or equal to C", "D>=C"
        GoTo ErrorHandler
    End If
    
    '=====================================
    ' CONSTRUCTION OF HexagonalSolid
    '=====================================
    Dim oGeomFactory As New ingrGeom3d.GeometryFactory
    Dim oPlane As ingrGeom3d.Plane3d
    Dim arrPoints(0 To 21) As Double
    Dim oPoint3d As ingrGeom3d.Point3d
    Dim line As Line3d
    
    
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d
    
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a, 0, 0)
    m_outputColl.AddOutput "Point2", oPoint3d
    
    arrPoints(0) = 0
    arrPoints(1) = -d / 2
    arrPoints(2) = -b / 2
    arrPoints(3) = 0
    arrPoints(4) = -c / 2
    arrPoints(5) = 0
    arrPoints(6) = 0
    arrPoints(7) = -d / 2
    arrPoints(8) = b / 2
    arrPoints(9) = 0
    arrPoints(10) = d / 2
    arrPoints(11) = b / 2
    arrPoints(12) = 0
    arrPoints(13) = c / 2
    arrPoints(14) = 0
    arrPoints(15) = 0
    arrPoints(16) = d / 2
    arrPoints(17) = -b / 2
    
    'Creation of Edge1 to 12
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
    m_outputColl.AddOutput "Edge1", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
    m_outputColl.AddOutput "Edge2", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
    m_outputColl.AddOutput "Edge3", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), arrPoints(12), arrPoints(13), arrPoints(14))
    m_outputColl.AddOutput "Edge4", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(12), arrPoints(13), arrPoints(14), arrPoints(15), arrPoints(16), arrPoints(17))
    m_outputColl.AddOutput "Edge5", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(15), arrPoints(16), arrPoints(17), arrPoints(0), arrPoints(1), arrPoints(2))
    m_outputColl.AddOutput "Edge6", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), a, arrPoints(1), arrPoints(2))
    m_outputColl.AddOutput "Edge7", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), a, arrPoints(4), arrPoints(5))
    m_outputColl.AddOutput "Edge8", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), a, arrPoints(7), arrPoints(8))
    m_outputColl.AddOutput "Edge9", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), a, arrPoints(10), arrPoints(11))
    m_outputColl.AddOutput "Edge10", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(12), arrPoints(13), arrPoints(14), a, arrPoints(13), arrPoints(14))
    m_outputColl.AddOutput "Edge11", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(15), arrPoints(16), arrPoints(17), a, arrPoints(16), arrPoints(17))
    m_outputColl.AddOutput "Edge12", line
    Set line = Nothing
'
'    'creation of points 3 to 8
'
'    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a / 2, (arrPoints(1) + arrPoints(4)) / 2, (arrPoints(2) + arrPoints(5)) / 2)
'    m_outputColl.AddOutput "Point3", oPoint3d
'
'    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a / 2, (arrPoints(4) + arrPoints(7)) / 2, (arrPoints(5) + arrPoints(8)) / 2)
'    m_outputColl.AddOutput "Point4", oPoint3d
'
'    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a / 2, (arrPoints(7) + arrPoints(10)) / 2, (arrPoints(8) + arrPoints(11)) / 2)
'    m_outputColl.AddOutput "Point5", oPoint3d
'
'    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a / 2, (arrPoints(10) + arrPoints(13)) / 2, (arrPoints(11) + arrPoints(14)) / 2)
'    m_outputColl.AddOutput "Point6", oPoint3d
'
'    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a / 2, (arrPoints(13) + arrPoints(16)) / 2, (arrPoints(14) + arrPoints(17)) / 2)
'    m_outputColl.AddOutput "Point7", oPoint3d
'
'    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, a / 2, (arrPoints(16) + arrPoints(1)) / 2, (arrPoints(17) + arrPoints(2)) / 2)
'    m_outputColl.AddOutput "Point8", oPoint3d
'
    'Creation of TopSurface
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 6, arrPoints)
    m_outputColl.AddOutput "TopSurface", oPlane
    
    arrPoints(0) = a
    arrPoints(1) = -d / 2
    arrPoints(2) = -b / 2
    arrPoints(3) = a
    arrPoints(4) = d / 2
    arrPoints(5) = -b / 2
    arrPoints(6) = a
    arrPoints(7) = c / 2
    arrPoints(8) = 0
    arrPoints(9) = a
    arrPoints(10) = d / 2
    arrPoints(11) = b / 2
    arrPoints(12) = a
    arrPoints(13) = -d / 2
    arrPoints(14) = b / 2
    arrPoints(15) = a
    arrPoints(16) = -c / 2
    arrPoints(17) = 0
    'Creation of BottomSurface
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 6, arrPoints)
    m_outputColl.AddOutput "BottomSurface", oPlane

    'Creation of Edge13 to 18
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
    m_outputColl.AddOutput "Edge13", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
    m_outputColl.AddOutput "Edge14", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
    m_outputColl.AddOutput "Edge15", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), arrPoints(12), arrPoints(13), arrPoints(14))
    m_outputColl.AddOutput "Edge16", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(12), arrPoints(13), arrPoints(14), arrPoints(15), arrPoints(16), arrPoints(17))
    m_outputColl.AddOutput "Edge17", line
    Set line = Nothing
    
    Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(15), arrPoints(16), arrPoints(17), arrPoints(0), arrPoints(1), arrPoints(2))
    m_outputColl.AddOutput "Edge18", line
    Set line = Nothing
    
    'Creation of surface1
    arrPoints(0) = 0
    arrPoints(1) = -d / 2
    arrPoints(2) = -b / 2
    arrPoints(3) = 0
    arrPoints(4) = d / 2
    arrPoints(5) = -b / 2
    arrPoints(6) = a
    arrPoints(7) = d / 2
    arrPoints(8) = -b / 2
    arrPoints(9) = a
    arrPoints(10) = -d / 2
    arrPoints(11) = -b / 2
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Surface1", oPlane

    'Creation of surface2
    arrPoints(0) = 0
    arrPoints(1) = -d / 2
    arrPoints(2) = b / 2
    arrPoints(3) = a
    arrPoints(4) = -d / 2
    arrPoints(5) = b / 2
    arrPoints(6) = a
    arrPoints(7) = d / 2
    arrPoints(8) = b / 2
    arrPoints(9) = 0
    arrPoints(10) = d / 2
    arrPoints(11) = b / 2
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Surface2", oPlane

    'Creation of surface3
    arrPoints(0) = 0
    arrPoints(1) = d / 2
    arrPoints(2) = -b / 2
    arrPoints(3) = 0
    arrPoints(4) = c / 2
    arrPoints(5) = 0
    arrPoints(6) = a
    arrPoints(7) = c / 2
    arrPoints(8) = 0
    arrPoints(9) = a
    arrPoints(10) = d / 2
    arrPoints(11) = -b / 2
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Surface3", oPlane

    'Creation of surface4
    arrPoints(0) = 0
    arrPoints(1) = -c / 2
    arrPoints(2) = 0
    arrPoints(3) = a
    arrPoints(4) = -c / 2
    arrPoints(5) = 0
    arrPoints(6) = a
    arrPoints(7) = -d / 2
    arrPoints(8) = b / 2
    arrPoints(9) = 0
    arrPoints(10) = -d / 2
    arrPoints(11) = b / 2
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Surface4", oPlane

    'Creation of surface5
    arrPoints(0) = 0
    arrPoints(1) = -c / 2
    arrPoints(2) = 0
    arrPoints(3) = 0
    arrPoints(4) = -d / 2
    arrPoints(5) = -b / 2
    arrPoints(6) = a
    arrPoints(7) = -d / 2
    arrPoints(8) = -b / 2
    arrPoints(9) = a
    arrPoints(10) = -c / 2
    arrPoints(11) = 0
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Surface5", oPlane
    
    'Creation of surface6
    arrPoints(0) = 0
    arrPoints(1) = d / 2
    arrPoints(2) = b / 2
    arrPoints(3) = a
    arrPoints(4) = d / 2
    arrPoints(5) = b / 2
    arrPoints(6) = a
    arrPoints(7) = c / 2
    arrPoints(8) = 0
    arrPoints(9) = 0
    arrPoints(10) = c / 2
    arrPoints(11) = 0
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Surface6", oPlane
    
    Set oPlane = Nothing
    Set oPoint3d = Nothing
    
     Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'Following method will be removed very soon. This has been
'included to improve the performance. These lines will be removed as soon as
'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.

' CMCache custom method to cache the input argument into a parameter contend and the reverse conversion
'It is up to you to find a way to convert your reference data object to a parameter content
Public Sub CMCacheForPart(pInputCM As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)

 If bArgToCache Then
       
        'Need to convert the graphic input pToConvert into a Parameter ( pOutput)
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        
        Dim oPart As IJDPart
        Set oPart = pToConvert
        'MsgBox "Partnumber :" & oPart.PartNumber
        
        ' Create a PC whose value is an identifier of the input
        ' Raju,
        ' the property Part_Number must be retrieved form from the pToConvert argument.
        ' I am hardcoding it
        '
        oPC.Type = igString
        oPC.String = oPart.PartNumber

        Set oPart = Nothing

        ' Always return this PC
        Set pOutput = oPC
        Set oPC = Nothing
 Else
        'Need to convert the cached Parameter pToConvert into your reference data object pOutput
        Dim oPCout As IJDParameterContent
        Dim oPCCache As IJDParameterContent
        Set oPCout = New DParameterContent
        
        ' Here there is three options
        ' o Return a parameter contents, that containts the part_number stored by the pToConvert argument
        '   In this case the edit command will have to retrieve the Part object when needed.
        '   Note : It is better to return a copy of the cached object instead the cached object itself.
        '          This allow to avoid the edition of the cached object.
        ' o Retrieve the Part_number and retrieve your part object with it.
        '   With this solution you can have assoc assertion while when this method is called you are in
        '   in a compute process.
        ' o Get the symbol (or equipment) from the pInputCM, query for the IJDReferencesArg interface,
        '   then get the argument at index 1 it is the SiteProxy.
        '   With this last method, the part has to be passed by reference to the symbol,
        '   but that is what you doing with your design.
        
        ' returning NULL means that the cache method is unable to resolve the cache.
        ' as of now anyway the part is passed as an input argument. this issue of caching
        ' has to be resolved in cycle2.
        
        Set pOutput = Nothing
        
        ' Here is the implementation of the option 1
'        Set oPCCache = pToConvert
'        oPCout.Type = oPCCache.Type
'        oPCout.String = oPCCache.String
'        Set pOutput = oPCout
'        Set oPCout = Nothing
        
        ' Here is the implementation of the option 2
'        Dim oRefDB As GSCADRefDataServices.IJDRefDBGlobal
'        Dim oPart As IJDPart

'        Set oRefDB = New GSCADRefDataServices.RefDBGlobal
'        Set oPart = oRefDB.GetCatalogPart("Storage Tanks", oPCCache.String)
'        Set pOutput = oRefDB
'        Set oRefDB = Nothing
'        Set oPart = Nothing
End If

End Sub

