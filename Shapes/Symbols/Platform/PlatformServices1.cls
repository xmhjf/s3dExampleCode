VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformServices1"
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
'   11 Feb MU Creation
'   30May2003   VS     Change the normal direction of the boundary imposed on plane so that
'                           the area inside the boundary is treated as material
'   24Jul2003   VS     Do not clear middle tier errors. This may clear out other errors too.
'   15Oct2003   VS     Added outputs for more edges. Also, Clean up the code and use
'                           CreateByOutrBdry instead of creating infinite planes
'   16Oct2003   VS     TR-CP·50735  Platform01 symbol should maitain the same output names.
'                           Earlier, at sweep angle 360, the top/bottom surface was being renamed to Plane1/Plane2
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private m_bAngle As Boolean
Dim m_outputColl As IJDOutputCollection
Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, _
                                      ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) _
                                                                                     As String
    IJDUserSymbolServices_GetDefinitionName = "Platform.PlatformServices1"
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
            Inputs(index).Description = "Length of the Cylinder"
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

    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of Platform 001"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
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
    O4.Name = "Revolution2"
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

    Dim O11 As IMSSymbolEntities.IJDOutput
    Set O11 = New IMSSymbolEntities.DOutput
    O11.Name = "Line5"
    O11.Description = "Line"
    O11.Properties = 0
    IJDOutputs.SetOutput O11
    Set O11 = Nothing

    Dim O12 As IMSSymbolEntities.IJDOutput
    Set O12 = New IMSSymbolEntities.DOutput
    O12.Name = "Line6"
    O12.Description = "Line"
    O12.Properties = 0
    IJDOutputs.SetOutput O12
    Set O12 = Nothing

    Dim O13 As IMSSymbolEntities.IJDOutput
    Set O13 = New IMSSymbolEntities.DOutput
    O13.Name = "Line7"
    O13.Description = "Line"
    O13.Properties = 0
    IJDOutputs.SetOutput O13
    Set O13 = Nothing

    Dim O14 As IMSSymbolEntities.IJDOutput
    Set O14 = New IMSSymbolEntities.DOutput
    O14.Name = "Line8"
    O14.Description = "Line"
    O14.Properties = 0
    IJDOutputs.SetOutput O14
    Set O14 = Nothing

    Dim O15 As IMSSymbolEntities.IJDOutput
    Set O15 = New IMSSymbolEntities.DOutput
    O15.Name = "Line9"
    O15.Description = "Line"
    O15.Properties = 0
    IJDOutputs.SetOutput O15
    Set O15 = Nothing

    Dim O16 As IMSSymbolEntities.IJDOutput
    Set O16 = New IMSSymbolEntities.DOutput
    O16.Name = "Line10"
    O16.Description = "Line"
    O16.Properties = 0
    IJDOutputs.SetOutput O16
    Set O16 = Nothing

    Dim O17 As IMSSymbolEntities.IJDOutput
    Set O17 = New IMSSymbolEntities.DOutput
    O17.Name = "Line11"
    O17.Description = "Line"
    O17.Properties = 0
    IJDOutputs.SetOutput O17
    Set O17 = Nothing

    Dim O18 As IMSSymbolEntities.IJDOutput
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
    PhysicalRepEval.ProgId = "Platform.PlatformServices1"
    
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
                                                ByVal definitionParameters As Variant, _
                                                ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "Platform.PlatformServices1"
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

Private Sub IJDUserSymbolServices_InvokeRepresentation( _
                                        ByVal pSymbolOccurrence As Object, _
                                        ByVal RepName As String, _
                                        ByVal OutputColl As Object, _
                                        arrayOfInputs() As Variant)
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
    Dim arrPoints(0 To 11) As Double
    Dim oGeomFactory       As GeometryFactory
    Dim oPlane             As Plane3d
    Dim objRevolution      As Object
    Dim oLine              As Line3d
    Dim oAxisVector        As IJDVector
    Dim oCenterPoint       As IJDPosition
    Dim oPoint3d           As Point3d
    Dim oComplexString     As ComplexString3d
    Dim oPart              As IJDPart
    
    Set oPart = arrayOfInputs(1)
    Set oGeomFactory = New GeometryFactory
    Set oAxisVector = New DVector
    Set oCenterPoint = New DPosition

    a = arrayOfInputs(2)
    b = arrayOfInputs(3)
    c = arrayOfInputs(4)
    d = arrayOfInputs(5)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If a <= 0 Or b <= 0 Or c <= 0 Or d <= 0 Then
        oErrors.Add E_FAIL, "Platform.PlatformServices1", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    End If
    
    'Case of 360 degrees.
    If (Abs(d - 6.283185) < TOLERANCE) Then
        m_bAngle = True
    Else
        While (d > 6.283185)
            d = d - 6.283185
            If (d < TOLERANCE) Then
                 m_bAngle = True
                 d = 6.2831
            End If
        Wend
    End If
    
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d
    Set oPoint3d = Nothing

    'Create a line and revolve it around the AxisVector with centre point and angle to get inner curved surface
    oCenterPoint.Set 0, 0, 0
    oAxisVector.Set 0, 0, 1
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, c * Sin(d / 2), c * Cos(d / 2), -b, c * Sin(d / 2), c * Cos(d / 2), 0)
    Set objRevolution = PlaceRevolution(Nothing, oLine, oAxisVector, oCenterPoint, d, False)
    Set oLine = Nothing
    m_outputColl.AddOutput "Revolution1", objRevolution
    Set objRevolution = Nothing
    
    'Create a line and revolve it around the AxisVector with centre point and angle to get outer curved surface
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), -b, (a + c) * Sin(d / 2), (a + c) * Cos(d / 2), 0)
    Set objRevolution = PlaceRevolution(Nothing, oLine, oAxisVector, oCenterPoint, d, False)
    Set oLine = Nothing
    m_outputColl.AddOutput "Revolution2", objRevolution
    Set objRevolution = Nothing
            
     If (m_bAngle = True) Then
            Dim oCircle As IngrGeom3D.Circle3d
            
            'Create the TopSurface and its edges
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, 0, 0, 0, 1, c + a)
            m_outputColl.AddOutput "Line4", oCircle
            Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oCircle)
            Set oCircle = Nothing
                        
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, 0, 0, 0, -1, c)
            m_outputColl.AddOutput "Line2", oCircle
            
            Set oComplexString = New ComplexString3d
            oComplexString.AddCurve oCircle, True
            oPlane.AddHole oComplexString
            m_outputColl.AddOutput "TopSurface", oPlane
            
            Set oCircle = Nothing
            Set oComplexString = Nothing
            Set oPlane = Nothing
            
            'Create the BottomSurface and its edges
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, -b, 0, 0, -1, c + a)
            m_outputColl.AddOutput "Line12", oCircle
            Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oCircle)
            Set oCircle = Nothing
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, -b, 0, 0, 1, c)
            m_outputColl.AddOutput "Line10", oCircle
            
            Set oComplexString = New ComplexString3d
            oComplexString.AddCurve oCircle, True
            oPlane.AddHole oComplexString
            m_outputColl.AddOutput "BottomSurface", oPlane
            
            Set oCircle = Nothing
            Set oComplexString = Nothing
            Set oPlane = Nothing
            
    Else 'If (m_bAngle = False) Then
        
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
            
        'Creation of Line5 and 6
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(9), arrPoints(10), arrPoints(11))
        m_outputColl.AddOutput "Line5", oLine
        Set oLine = Nothing
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
        m_outputColl.AddOutput "Line6", oLine
        Set oLine = Nothing

        'Creates a plane based on the above points
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
        
        'Creation of Edge4, 5 and 6
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

        'Create the persistent curves and plane for the topsurface of the platform
        Set oComplexString = CreateTopCurves(Nothing, oCenterPoint, c, d, a, b)
        Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
        m_outputColl.AddOutput "TopSurface", oPlane
        Set oPlane = Nothing
        Set oComplexString = Nothing
        
        'Create the persistent curves and plane for the bottomsurface of the platform
        oCenterPoint.Set 0, 0, -b
        Set oComplexString = CreateBottomCurves(Nothing, oCenterPoint, c, d, a, b)
        Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
        m_outputColl.AddOutput "BottomSurface", oPlane
        Set oPlane = Nothing
        Set oComplexString = Nothing
    End If
    
    Set oAxisVector = Nothing
    Set oCenterPoint = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'''This function creates the persistent curves based on the points provided
'''<{(CreateTopCurves Begin)}>
Private Function CreateTopCurves(ByVal oResourceManager As IUnknown, _
                            ByVal oCenterPoint As IJDPosition, _
                            ByVal c As Double, _
                            ByVal d As Double, _
                            ByVal a As Double, _
                            ByVal b As Double) As IngrGeom3D.ComplexString3d
 
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As Object
    Dim oCircle         As IngrGeom3D.Circle3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim arrPt(12)       As New AutoMath.DPosition
    Dim oNormalVector   As IJDVector
    Dim oLines          As Collection
    
    On Error GoTo ErrorHandler
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oNormalVector = New AutoMath.DVector
    Set oLines = New Collection
    
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
    
    Set oLine = PlaceLine(oResourceManager, arrPt(1), arrPt(2))
    oLines.Add oLine
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line1", oLine
    End If
    oNormalVector.Set 0, 0, 1
    
    Set oArc = PlaceArcByCen(oResourceManager, oCenterPoint, arrPt(2), arrPt(3), oNormalVector)
    oLines.Add oArc
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line2", oArc
    End If
    Set oLine = PlaceLine(oResourceManager, arrPt(3), arrPt(4))
    oLines.Add oLine
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line3", oLine
    End If
    oNormalVector.Set 0, 0, -1
    Set oArc = PlaceArcByCen(oResourceManager, oCenterPoint, arrPt(4), arrPt(1), oNormalVector)
    oLines.Add oArc
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line4", oArc
    End If
    Set CreateTopCurves = PlaceCString(arrPt(1), oLines)
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To oLines.count
        oLines.Remove 1
    Next iCount
    Set oLines = Nothing
    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(CreateTopCurves End)}>

'''This function creates the persistent curves based on the points provided
'''<{(CreateBottomCurves Begin)}>
Private Function CreateBottomCurves(ByVal oResourceManager As IUnknown, _
                            ByVal oCenterPoint As IJDPosition, _
                            ByVal c As Double, _
                            ByVal d As Double, _
                            ByVal a As Double, _
                            ByVal b As Double) As IngrGeom3D.ComplexString3d
 
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As Object
    Dim oCircle         As IngrGeom3D.Circle3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim arrPt(12)       As New AutoMath.DPosition
    Dim oNormalVector   As IJDVector
    Dim oLines          As Collection
    
    On Error GoTo ErrorHandler
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oNormalVector = New AutoMath.DVector
    Set oLines = New Collection
    
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
    
    Set oLine = PlaceLine(oResourceManager, arrPt(1), arrPt(2))
    oLines.Add oLine
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line9", oLine
    End If
    oNormalVector.Set 0, 0, -1
    
    Set oArc = PlaceArcByCen(oResourceManager, oCenterPoint, arrPt(2), arrPt(3), oNormalVector)
    oLines.Add oArc
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line10", oArc
    End If
    Set oLine = PlaceLine(oResourceManager, arrPt(3), arrPt(4))
    oLines.Add oLine
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line11", oLine
    End If
    oNormalVector.Set 0, 0, 1
    Set oArc = PlaceArcByCen(oResourceManager, oCenterPoint, arrPt(4), arrPt(1), oNormalVector)
    oLines.Add oArc
    If Not oResourceManager Is Nothing Then
        m_outputColl.AddOutput "Line12", oArc
    End If
    Set CreateBottomCurves = PlaceCString(arrPt(1), oLines)
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To oLines.count
        oLines.Remove 1
    Next iCount
    Set oLines = Nothing
    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(CreateBottomCurves End)}>
