VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformServices2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    PlatformServices.cls
'
'Author
'    MU
'
'Description
'    Definition of PlatformServices Symbol
'History
'   11 Feb Mule Creation
'   30May2003   VS     Change the normal direction of the boundary imposed on plane so that
'                           the area inside the boundary is treated as material
'   24Jul2003   VS      Do not clear middle tier errors. This may clear out other errors too.
'   15Oct2003   VS      CR36485 - Added outputs for more edges. Also, clean up the code a bit
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection
Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, _
                                     ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName( _
                                  ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "Platform.PlatformServices2"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition( _
                            pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
                            
    On Error GoTo ErrorHandler
   
   ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    'Feed Platform Defnition
    ' Inputs:
    '          1. "Part"
    '          2. "PlatformWidth"
    '          3. "PlatformHeight"
    '          4. "PlatformRadius"
    '          5. "PlatformAngle"
    ' Representations :
    '           Physical
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 5) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'Define inputs for Platform
    Dim index As Integer
    For index = 1 To 4
        Set Inputs(index) = New IMSSymbolEntities.DInput
        Inputs(index).Properties = igINPUT_IS_A_PARAMETER
        If (index = 1) Then
            PC.UomValue = 0.2
            Inputs(index).Name = "PlatformWidth"
            Inputs(index).Description = "Platform Width"
        ElseIf (index = 2) Then
            PC.UomValue = 0.2
            Inputs(index).Name = "PlatformHeight"
        ElseIf (index = 3) Then
            PC.UomValue = 0.08
            Inputs(index).Name = "PlatformRadius"
        ElseIf (index = 4) Then
            PC.UomValue = PI / 4
            Inputs(index).Name = "PlatformAngle"
        End If
        Inputs(index).DefaultParameterValue = PC
    Next
    
    For index = 1 To 4
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next
    
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set rep1 = New IMSSymbolEntities.DRepresentation
    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of Platform002"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical
    Set IJDOutputs = rep1

    'Define and set the outputs for this representation
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput
    O1.Name = "Plane1"
    O1.Description = "Plane"
    O1.Properties = 0
    IJDOutputs.SetOutput O1
    Set O1 = Nothing
    
    Dim O2 As IMSSymbolEntities.DOutput
    Set O2 = New IMSSymbolEntities.DOutput
    O2.Name = "Plane2"
    O2.Description = "Plane"
    O2.Properties = 0
    IJDOutputs.SetOutput O2
    Set O2 = Nothing
    
    Dim O3 As IMSSymbolEntities.DOutput
    Set O3 = New IMSSymbolEntities.DOutput
    O3.Name = "Revolution1"
    O3.Description = "Plane"
    O3.Properties = 0
    IJDOutputs.SetOutput O3
    Set O3 = Nothing
    
    Dim O4 As IMSSymbolEntities.DOutput
    Set O4 = New IMSSymbolEntities.DOutput
    O4.Name = "Plane3"
    O4.Description = "Plane"
    O4.Properties = 0
    IJDOutputs.SetOutput O4
    Set O4 = Nothing
    
    Dim O5 As IMSSymbolEntities.DOutput
    Set O5 = New IMSSymbolEntities.DOutput
    O5.Name = "TopSurface"
    O5.Description = "Plane"
    O5.Properties = 0
    IJDOutputs.SetOutput O5
    Set O5 = Nothing
    
    Dim O6 As IMSSymbolEntities.DOutput
    Set O6 = New IMSSymbolEntities.DOutput
    O6.Name = "BottomSurface"
    O6.Description = "Plane"
    O6.Properties = 0
    IJDOutputs.SetOutput O6
    Set O6 = Nothing
    
    Dim O7 As IMSSymbolEntities.DOutput
    Set O7 = New IMSSymbolEntities.DOutput
    O7.Name = "Line1"
    O7.Description = "Line"
    O7.Properties = 0
    IJDOutputs.SetOutput O7
    Set O7 = Nothing
    
    Dim O8 As IMSSymbolEntities.DOutput
    Set O8 = New IMSSymbolEntities.DOutput
    O8.Name = "Line2"
    O8.Description = "Line"
    O8.Properties = 0
    IJDOutputs.SetOutput O8
    Set O8 = Nothing
    
    Dim O9 As IMSSymbolEntities.DOutput
    Set O9 = New IMSSymbolEntities.DOutput
    O9.Name = "Line3"
    O9.Description = "Line"
    O9.Properties = 0
    IJDOutputs.SetOutput O9
    Set O9 = Nothing
    
    Dim O10 As IMSSymbolEntities.DOutput
    Set O10 = New IMSSymbolEntities.DOutput
    O10.Name = "Line4"
    O10.Description = "Line"
    O10.Properties = 0
    IJDOutputs.SetOutput O10
    Set O10 = Nothing

    Dim O11 As IMSSymbolEntities.DOutput
    Set O11 = New IMSSymbolEntities.DOutput
    O11.Name = "Line5"
    O11.Description = "Line"
    O11.Properties = 0
    IJDOutputs.SetOutput O11
    Set O11 = Nothing

    Dim O12 As IMSSymbolEntities.DOutput
    Set O12 = New IMSSymbolEntities.DOutput
    O12.Name = "Line6"
    O12.Description = "Line"
    O12.Properties = 0
    IJDOutputs.SetOutput O12
    Set O12 = Nothing

    Dim O13 As IMSSymbolEntities.DOutput
    Set O13 = New IMSSymbolEntities.DOutput
    O13.Name = "Line7"
    O13.Description = "Line"
    O13.Properties = 0
    IJDOutputs.SetOutput O13
    Set O13 = Nothing

    Dim O14 As IMSSymbolEntities.DOutput
    Set O14 = New IMSSymbolEntities.DOutput
    O14.Name = "Line8"
    O14.Description = "Line"
    O14.Properties = 0
    IJDOutputs.SetOutput O14
    Set O14 = Nothing
    
    Dim O15 As IMSSymbolEntities.DOutput
    Set O15 = New IMSSymbolEntities.DOutput
    O15.Name = "Line9"
    O15.Description = "Line"
    O15.Properties = 0
    IJDOutputs.SetOutput O15
    Set O15 = Nothing
    
    Dim O16 As IMSSymbolEntities.DOutput
    Set O16 = New IMSSymbolEntities.DOutput
    O16.Name = "Line10"
    O16.Description = "Line"
    O16.Properties = 0
    IJDOutputs.SetOutput O16
    Set O16 = Nothing
    
    Dim O17 As IMSSymbolEntities.DOutput
    Set O17 = New IMSSymbolEntities.DOutput
    O17.Name = "Line11"
    O17.Description = "Line"
    O17.Properties = 0
    IJDOutputs.SetOutput O17
    Set O17 = Nothing
    
    Dim O18 As IMSSymbolEntities.DOutput
    Set O18 = New IMSSymbolEntities.DOutput
    O18.Name = "Line12"
    O18.Description = "Line"
    O18.Properties = 0
    IJDOutputs.SetOutput O18
    Set O18 = Nothing

    Dim oPoint1 As IMSSymbolEntities.DOutput
    Set oPoint1 = New IMSSymbolEntities.DOutput
    oPoint1.Name = "Point1"
    oPoint1.Description = "Platform"
    oPoint1.Properties = 0
    IJDOutputs.SetOutput oPoint1
    Set oPoint1 = Nothing
    
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
    PhysicalRepEval.Description = "Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "Platform.PlatformServices2"
    
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

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
      ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "Platform.PlatformServices2"
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
'CREATION OF PHYSICAL REPRESENTATION OF PLATFORM
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    'The planes are constrcuted in such a way that normals point outwards
    On Error GoTo ErrorHandler
    Dim a As Double, b As Double, c As Double, d As Double
    Dim arrPoints(0 To 11)  As Double
    Dim oGeomFactory     As GeometryFactory
    Dim oPlane           As Plane3d
    Dim objRevolution    As Object
    Dim oLine            As Line3d
    Dim oAxisVector      As IJDVector
    Dim oCenterPoint     As IJDPosition
    Dim oComplexString   As ComplexString3d
    Dim oPoint3d         As Point3d
    Dim oPart            As IJDPart
    
    Set oPart = arrayOfInputs(1)
    a = arrayOfInputs(2)
    b = arrayOfInputs(3)
    c = arrayOfInputs(4)
    d = arrayOfInputs(5)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
   
    If a <= 0 Or b <= 0 Or c <= 0 Or d <= 0 Then
        oErrors.Add E_FAIL, "Platform.PlatformServices2", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    End If
    
    'Case of 180 degrees.
    If (Abs(d - PI) < TOLERANCE) Then
        ' do nothing
    Else
        While (d > PI)
            d = d - PI
            If (d < TOLERANCE) Then
                d = PI
            End If
        Wend
    End If
    
    If d = PI Then
        oErrors.Add E_FAIL, "Platform.PlatformServices2", "Angle must be less than 180 degrees", "D=PI"
        GoTo ErrorHandler
    ElseIf ((c * (1 - Cos(d / 2))) >= (a * Cos(d / 2))) Then
        oErrors.Add E_FAIL, "Platform.PlatformServices2", "Shape Dimensions are not valid", "NotValid"
        GoTo ErrorHandler
    End If
    
    arrPoints(0) = -c * Sin(d / 2)
    arrPoints(1) = c * Cos(d / 2)
    arrPoints(2) = 0
    arrPoints(3) = -(a + c) * Sin(d / 2)
    arrPoints(4) = (a + c) * Cos(d / 2)
    arrPoints(5) = 0
    arrPoints(6) = -(a + c) * Sin(d / 2)
    arrPoints(7) = (a + c) * Cos(d / 2)
    arrPoints(8) = -b
    arrPoints(9) = -c * Sin(d / 2)
    arrPoints(10) = c * Cos(d / 2)
    arrPoints(11) = -b
    
    'Creates a plane based on the above points
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d
    Set oPoint3d = Nothing

    'Creation of Line5 and 6
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(9), arrPoints(10), arrPoints(11))
    m_outputColl.AddOutput "Line5", oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
    m_outputColl.AddOutput "Line6", oLine
    Set oLine = Nothing

    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Plane1", oPlane
    Set oPlane = Nothing
     
    arrPoints(0) = c * Sin(d / 2)
    arrPoints(1) = c * Cos(d / 2)
    arrPoints(2) = 0
    arrPoints(3) = c * Sin(d / 2)
    arrPoints(4) = c * Cos(d / 2)
    arrPoints(5) = -b
    arrPoints(6) = (a + c) * Sin(d / 2)
    arrPoints(7) = (a + c) * Cos(d / 2)
    arrPoints(8) = -b
    arrPoints(9) = (a + c) * Sin(d / 2)
    arrPoints(10) = (a + c) * Cos(d / 2)
    arrPoints(11) = 0

    'Creation of Line7 and 8
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
    m_outputColl.AddOutput "Line7", oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
    m_outputColl.AddOutput "Line8", oLine
    Set oLine = Nothing

    'Creates a plane based on the above points
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Plane2", oPlane
    Set oPlane = Nothing
    
    arrPoints(0) = -(a + c) * Sin(d / 2)
    arrPoints(1) = (a + c) * Cos(d / 2)
    arrPoints(2) = 0
    arrPoints(3) = (a + c) * Sin(d / 2)
    arrPoints(4) = (a + c) * Cos(d / 2)
    arrPoints(5) = 0
    arrPoints(6) = (a + c) * Sin(d / 2)
    arrPoints(7) = (a + c) * Cos(d / 2)
    arrPoints(8) = -b
    arrPoints(9) = -(a + c) * Sin(d / 2)
    arrPoints(10) = (a + c) * Cos(d / 2)
    arrPoints(11) = -b
    Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
    m_outputColl.AddOutput "Plane3", oPlane
    Set oPlane = Nothing
    
    'Create a line and revolve it around the AxisVector with centre point and angle
    Set oAxisVector = New AutoMath.DVector
    Set oCenterPoint = New AutoMath.DPosition
    oCenterPoint.Set 0, 0, 0
    oAxisVector.Set 0, 0, 1
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, c * Sin(d / 2), c * Cos(d / 2), -b, c * Sin(d / 2), c * Cos(d / 2), 0)
    
    Set objRevolution = PlaceRevolution(Nothing, oLine, oAxisVector, oCenterPoint, d, True)
    Set oLine = Nothing
    m_outputColl.AddOutput "Revolution1", objRevolution
    Set objRevolution = Nothing
    
    'Create a infinite plane and impose boundaries
    ' based on the complexstring of curves
    oCenterPoint.Set 0, 0, -b
    Set oComplexString = CreateBottomRectangle(oCenterPoint, c, d, a, b)
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
    m_outputColl.AddOutput "BottomSurface", oPlane
    Set oPlane = Nothing
    Set oComplexString = Nothing
    
    'Create a infinite plane and impose boundaries
    'based on the complexstring of curves
    oCenterPoint.Set 0, 0, 0
    Set oComplexString = CreateTopRectangle(oCenterPoint, c, d, a, b)
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
    m_outputColl.AddOutput "TopSurface", oPlane
    Set oPlane = Nothing
    Set oComplexString = Nothing
    
    'Create the persistent curves on the topsurface of the platform
    oCenterPoint.Set 0, 0, 0
    CreateCurves oCenterPoint, c, d, a, b 'Top Surface
    oCenterPoint.Set 0, 0, -b
    CreateCurves oCenterPoint, c, d, a, b 'Bottom Surface
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'''This Function Creates a ComplexString based on the curves provided
'''<{(CreateTopRectangle begin)}>
Private Function CreateTopRectangle(ByVal oCenterPoint As IJDPosition, _
                            ByVal c As Double, _
                            ByVal d As Double, _
                            ByVal a As Double, _
                            ByVal b As Double) As ComplexString3d
 
    Dim oLines          As Collection
    Dim oLine           As Line3d
    Dim oArc            As Arc3d
    Dim oGeomFactory    As GeometryFactory
    Dim objCStr         As Object
    Dim arrPt(12)       As New DPosition
    Dim oNormalVector   As IJDVector
    
    On Error GoTo ErrorHandler
    
    Set oLines = New Collection
    Set oGeomFactory = New GeometryFactory
    Set oNormalVector = New DVector
    
    If (oCenterPoint.z = 0) Then
        arrPt(4).Set -c * Sin(d / 2), c * Cos(d / 2), 0
        arrPt(3).Set -(a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0
        arrPt(2).Set (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0
        arrPt(1).Set c * Sin(d / 2), c * Cos(d / 2), 0
    Else
        arrPt(4).Set -c * Sin(d / 2), c * Cos(d / 2), -b
        arrPt(3).Set -(a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b
        arrPt(2).Set (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b
        arrPt(1).Set c * Sin(d / 2), c * Cos(d / 2), -b
    End If
    
    oNormalVector.Set 0, 0, -1
    Set oLine = PlaceLine(Nothing, arrPt(1), arrPt(2))
    oLines.Add oLine
    Set oLine = PlaceLine(Nothing, arrPt(2), arrPt(3))
    oLines.Add oLine
    Set oLine = PlaceLine(Nothing, arrPt(3), arrPt(4))
    oLines.Add oLine
    Set oArc = PlaceArcByCen(Nothing, oCenterPoint, arrPt(4), arrPt(1), oNormalVector)
    oLines.Add oArc
    Set objCStr = PlaceCString(arrPt(1), oLines)
    Set CreateTopRectangle = objCStr
    Set oLine = Nothing
    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(CreateTopRectangle end)}>

'''This Function Creates Persistent individual Curves
'''for the Top and Bottom Surface of the Platform
'''<{(CreateCurves begin)}>
Private Function CreateCurves(ByVal oCenterPoint As IJDPosition, _
                            ByVal c As Double, _
                            ByVal d As Double, _
                            ByVal a As Double, _
                            ByVal b As Double)
 
    Dim oLine1          As Line3d
    Dim oLine2          As Line3d
    Dim oLine3          As Line3d
    Dim oArc            As Object
    Dim oGeomFactory    As GeometryFactory
    Dim arrPt(12)       As New DPosition
    Dim oNormalVector   As IJDVector
    
    On Error GoTo ErrorHandler
    
    Set oGeomFactory = New GeometryFactory
    Set oNormalVector = New DVector
    
    If (oCenterPoint.z = 0) Then
        arrPt(1).Set -c * Sin(d / 2), c * Cos(d / 2), 0
        arrPt(2).Set -(a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0
        arrPt(3).Set (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0
        arrPt(4).Set c * Sin(d / 2), c * Cos(d / 2), 0
    Else
        arrPt(1).Set -c * Sin(d / 2), c * Cos(d / 2), -b
        arrPt(2).Set -(a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b
        arrPt(3).Set (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b
        arrPt(4).Set c * Sin(d / 2), c * Cos(d / 2), -b
    End If
    
    oNormalVector.Set 0, 0, 1
    
    Set oArc = PlaceArcByCen(Nothing, oCenterPoint, arrPt(4), arrPt(1), oNormalVector)
    Set oLine1 = PlaceLine(Nothing, arrPt(1), arrPt(2))
    Set oLine2 = PlaceLine(Nothing, arrPt(2), arrPt(3))
    Set oLine3 = PlaceLine(Nothing, arrPt(3), arrPt(4))
    
    If (oCenterPoint.z = 0) Then
        m_outputColl.AddOutput "Line1", oLine1
        m_outputColl.AddOutput "Line2", oLine2
        m_outputColl.AddOutput "Line3", oLine3
        m_outputColl.AddOutput "Line4", oArc
    Else
        m_outputColl.AddOutput "Line9", oLine1
        m_outputColl.AddOutput "Line10", oLine2
        m_outputColl.AddOutput "Line11", oLine3
        m_outputColl.AddOutput "Line12", oArc
    End If
    
    
    Set oLine1 = Nothing
    Set oLine2 = Nothing
    Set oLine3 = Nothing
    Set oArc = Nothing
    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(CreateCurves end)}>



'''This Function Creates a ComplexString based on the curves provided
'''<{(CreateBottomRectangle begin)}>
Private Function CreateBottomRectangle(ByVal oCenterPoint As IJDPosition, _
                            ByVal c As Double, _
                            ByVal d As Double, _
                            ByVal a As Double, _
                            ByVal b As Double) As ComplexString3d
 
    Dim oLines          As Collection
    Dim oLine           As Line3d
    Dim oArc            As Arc3d
    Dim oGeomFactory    As GeometryFactory
    Dim objCStr         As Object
    Dim arrPt(12)       As New DPosition
    Dim oNormalVector   As IJDVector
    
    On Error GoTo ErrorHandler
    
    Set oLines = New Collection
    Set oGeomFactory = New GeometryFactory
    Set oNormalVector = New DVector
    
    If (oCenterPoint.z = 0) Then
        arrPt(1).Set -c * Sin(d / 2), c * Cos(d / 2), 0
        arrPt(2).Set -(a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0
        arrPt(3).Set (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0
        arrPt(4).Set c * Sin(d / 2), c * Cos(d / 2), 0
    Else
        arrPt(1).Set -c * Sin(d / 2), c * Cos(d / 2), -b
        arrPt(2).Set -(a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b
        arrPt(3).Set (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b
        arrPt(4).Set c * Sin(d / 2), c * Cos(d / 2), -b
    End If
    
    oNormalVector.Set 0, 0, 1
    Set oLine = PlaceLine(Nothing, arrPt(1), arrPt(2))
    oLines.Add oLine
    Set oLine = PlaceLine(Nothing, arrPt(2), arrPt(3))
    oLines.Add oLine
    Set oLine = PlaceLine(Nothing, arrPt(3), arrPt(4))
    oLines.Add oLine
    Set oArc = PlaceArcByCen(Nothing, oCenterPoint, arrPt(4), arrPt(1), oNormalVector)
    oLines.Add oArc
    Set objCStr = PlaceCString(arrPt(1), oLines)
    Set CreateBottomRectangle = objCStr
    Set oLine = Nothing
    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(CreateBottomRectangle end)}>
