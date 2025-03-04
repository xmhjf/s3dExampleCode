VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectangularToriServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    RectangularToriServices.cls
'
'Author
'    MU
'
'Description
'    Definition of RectangularToriServices Symbol
'History
'10 Apr 2002    MU       Creation
'30 May 2003    VS     Change the normal direction of the boundary imposed on plane so that
'                           the area inside the boundary is treated as material
'02 Jun 2003    VS      Change the creation of end planes to use the CreateByOuterBdry method,
'                           this allows to check if the boundary has been created properly
'09 Jul 2003    SymbolTeam  Copyright Information, Header  is added/Updated.
'               (India)
'24 Jul 2003    VS      Do not clear middle tier errors. This may clear out other errors too.
'16 Oct 2003    VS      Added edges at the faces for CR36485. SmartSketch will locate points on these edges
'09 Nov 2004    VS      TR64758 - Change the symbol so that the X axis is perpendicular to the face.
'
'Notes
'
'    <notes>
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  19.DEC.2007     PK            Undone the changes made in TR-132021
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const E_FAIL = -2147467259
Private Const PI = 3.14159265
Private Const TOLERANCE = 0.000001
Private m_bAngle As Boolean 'Flag required to take care of 360 degrees
Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "RectangularTori.RectangularToriServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    ' Feed RectangularTori Definition
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
    
    'Define inputs for RectangularTori
    Dim Index As Integer
    For Index = 1 To 4
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.08
            Inputs(Index).Name = "A"
        ElseIf (Index = 2) Then
            PC.UomValue = 0.08
            Inputs(Index).Name = "B"
        ElseIf (Index = 3) Then
            PC.UomValue = 0.08
            Inputs(Index).Name = "C"
        ElseIf (Index = 4) Then
            PC.UomValue = PI / 4
            Inputs(Index).Name = "D"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 4
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next
 
    'Define the outputs
    Dim oSurface1 As IMSSymbolEntities.DOutput
    Set oSurface1 = New IMSSymbolEntities.DOutput
    
    oSurface1.Name = "Surface1"
    oSurface1.Description = "Surface1 of RectangularTori"
    oSurface1.Properties = 0
    
    Dim oSurface2 As IMSSymbolEntities.DOutput
    Set oSurface2 = New IMSSymbolEntities.DOutput
    
    oSurface2.Name = "Surface2"
    oSurface2.Description = "Surface2 of RectangularTori"
    oSurface2.Properties = 0
    
    Dim oSurface3 As IMSSymbolEntities.DOutput
    Set oSurface3 = New IMSSymbolEntities.DOutput
    
    oSurface3.Name = "Surface3"
    oSurface3.Description = "Surface3 of RectangularTori"
    oSurface3.Properties = 0
    
    Dim oSurface4 As IMSSymbolEntities.DOutput
    Set oSurface4 = New IMSSymbolEntities.DOutput
    
    oSurface4.Name = "Surface4"
    oSurface4.Description = "Surface4 of RectangularTori"
    oSurface4.Properties = 0
    
    Dim oSurface5 As IMSSymbolEntities.DOutput
    Set oSurface5 = New IMSSymbolEntities.DOutput
    
    oSurface5.Name = "Surface5"
    oSurface5.Description = "Surface5 of RectangularTori"
    oSurface5.Properties = 0
    
    Dim oSurface6 As IMSSymbolEntities.DOutput
    Set oSurface6 = New IMSSymbolEntities.DOutput
    
    oSurface6.Name = "Surface6"
    oSurface6.Description = "Surface6 of RectangularTori"
    oSurface6.Properties = 0
   
    Dim oPoint1 As IMSSymbolEntities.DOutput
    Set oPoint1 = New IMSSymbolEntities.DOutput
    
    oPoint1.Name = "Point1"
    oPoint1.Description = "RectangularTorus"
    oPoint1.Properties = 0
    
    Dim oPoint2 As IMSSymbolEntities.DOutput
    Set oPoint2 = New IMSSymbolEntities.DOutput
    
    oPoint2.Name = "Point2"
    oPoint2.Description = "RectangularTorus"
    oPoint2.Properties = 0
    
    
    Dim oPoint3 As IMSSymbolEntities.DOutput
    Set oPoint3 = New IMSSymbolEntities.DOutput
    
    oPoint3.Name = "Point3"
    oPoint3.Description = "RectangularTorus"
    oPoint3.Properties = 0
    
    Dim E1 As IMSSymbolEntities.DOutput
    Set E1 = New IMSSymbolEntities.DOutput
    E1.Name = "Edge1"
    E1.Description = "Edge1 of RectangularTorus"
    E1.Properties = 0
    
    Dim E2 As IMSSymbolEntities.DOutput
    Set E2 = New IMSSymbolEntities.DOutput
    E2.Name = "Edge2"
    E2.Description = "Edge2 of RectangularTorus"
    E2.Properties = 0
    
    Dim E3 As IMSSymbolEntities.DOutput
    Set E3 = New IMSSymbolEntities.DOutput
    E3.Name = "Edge3"
    E3.Description = "Edge3 of RectangularTorus"
    E3.Properties = 0
    
    Dim E4 As IMSSymbolEntities.DOutput
    Set E4 = New IMSSymbolEntities.DOutput
    E4.Name = "Edge4"
    E4.Description = "Edge4 of RectangularTorus"
    E4.Properties = 0
    
    Dim E5 As IMSSymbolEntities.DOutput
    Set E5 = New IMSSymbolEntities.DOutput
    E5.Name = "Edge5"
    E5.Description = "Edge5 of RectangularTorus"
    E5.Properties = 0
    
    Dim E6 As IMSSymbolEntities.DOutput
    Set E6 = New IMSSymbolEntities.DOutput
    E6.Name = "Edge6"
    E6.Description = "Edge6 of RectangularTorus"
    E6.Properties = 0
    
    Dim E7 As IMSSymbolEntities.DOutput
    Set E7 = New IMSSymbolEntities.DOutput
    E7.Name = "Edge7"
    E7.Description = "Edge7 of RectangularTorus"
    E7.Properties = 0
    
    Dim E8 As IMSSymbolEntities.DOutput
    Set E8 = New IMSSymbolEntities.DOutput
    E8.Name = "Edge8"
    E8.Description = "Edge8 of RectangularTorus"
    E8.Properties = 0
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of RectangularTori"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1
    
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
    
    IJDOutputs.SetOutput oPoint3
    Set oPoint3 = Nothing
    
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
    PhysicalRepEval.Description = "Physical representation of RectangularTori"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "RectangularTori.RectangularToriServices"
    
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
    oSymbolDefinition.ProgId = "RectangularTori.RectangularToriServices"
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
'CREATION OF PHYSICAL REPRESENTATION OF RectangularTori
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())

'       The points on the torus are numbered as below, 1-4 are
'       points on the bottom plane ans 5-8 are points on the top plane
'       2,6
'                   3,7
'       1,5
'
'                   4,8

    On Error GoTo ErrorHandler
    Dim a  As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    
    a = arrayOfInputs(2)
    b = arrayOfInputs(3)
    c = arrayOfInputs(4)
    d = arrayOfInputs(5)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If a <= 0 Or b <= 0 Or c <= 0 Or d <= 0 Then
        oErrors.Add E_FAIL, "RectangularTori.RectangularToriServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    ElseIf a >= 2 * c Then
        oErrors.Add E_FAIL, "RectangularTori.RectangularToriServices", "A should not be greater than twice C", "A>2*C"
        GoTo ErrorHandler
    End If
    
    'Case of 360 (or multiples of 360) degrees.
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
    

    '=====================================
    ' CONSTRUCTION OF RectangularTori
    '=====================================
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oPlane As IngrGeom3D.Plane3d
    Dim arrPoints(0 To 11) As Double
    Dim oUpVector As IJDVector
    Dim oDownVector As IJDVector
    Dim oCenterPoint As IJDPosition
    
    Dim oPoint1 As IJDPosition
    Dim oPoint2 As IJDPosition
    Dim oPoint3 As IJDPosition
    Dim oPoint4 As IJDPosition
    Dim oPoint5 As IJDPosition
    Dim oPoint6 As IJDPosition
    Dim oPoint7 As IJDPosition
    Dim oPoint8 As IJDPosition
    
    Dim oLine  As Line3d
    Dim oPoint3d As IngrGeom3D.Point3d
    Dim oObjRevolution As Object
    Dim oComplexString As IngrGeom3D.ComplexString3d
    Dim line As Line3d
    
    Dim dCentreX As Double, dCentreY As Double, dCentreZ As Double
    Dim dTopPlaneZ As Double, dBottomPlaneZ As Double
    Dim dOuterRadius As Double, dInnerRadius As Double
     
    'Initialize the vertices , the centre point and the axis of revolution
    dCentreX = 0
    dCentreY = -c
    dOuterRadius = c + a / 2
    dInnerRadius = c - a / 2
    dTopPlaneZ = b / 2
    dBottomPlaneZ = -b / 2
    
    Set oUpVector = New DVector
    Set oDownVector = New DVector
    Set oCenterPoint = New DPosition
    Set oPoint1 = New DPosition
    Set oPoint2 = New DPosition
    Set oPoint3 = New DPosition
    Set oPoint4 = New DPosition
    Set oPoint5 = New DPosition
    Set oPoint6 = New DPosition
    Set oPoint7 = New DPosition
    Set oPoint8 = New DPosition
    
    oCenterPoint.Set dCentreX, dCentreY, 0
    oUpVector.Set 0, 0, 1
    oDownVector.Set 0, 0, -1
    
    oPoint1.Set 0, -a / 2, dBottomPlaneZ
    oPoint2.Set 0, a / 2, dBottomPlaneZ
    oPoint3.Set (dOuterRadius * Sin(d)), ((dOuterRadius * Cos(d)) - c), dBottomPlaneZ
    oPoint4.Set (dInnerRadius * Sin(d)), ((dInnerRadius * Cos(d)) - c), dBottomPlaneZ
    oPoint5.Set 0, -a / 2, dTopPlaneZ
    oPoint6.Set 0, a / 2, dTopPlaneZ
    oPoint7.Set (dOuterRadius * Sin(d)), ((dOuterRadius * Cos(d)) - c), dTopPlaneZ
    oPoint8.Set (dInnerRadius * Sin(d)), ((dInnerRadius * Cos(d)) - c), dTopPlaneZ
    
    'Creation of centre point
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, oCenterPoint.x, oCenterPoint.y, oCenterPoint.z)
    m_outputColl.AddOutput "Point3", oPoint3d
    Set oPoint3d = Nothing
    
    'Creation of Surface1
    Set oLine = PlaceLine(Nothing, oPoint1, oPoint5)
    Set oObjRevolution = PlaceRevolution(Nothing, oLine, oDownVector, _
                                        oCenterPoint, d, True)
    m_outputColl.AddOutput "Surface1", oObjRevolution
    Set oObjRevolution = Nothing
    
    'Creation of Surface2
    Set oLine = PlaceLine(Nothing, oPoint2, oPoint6)
    Set oObjRevolution = PlaceRevolution(Nothing, oLine, oDownVector, _
                                        oCenterPoint, d, True)
    Set oLine = Nothing
    m_outputColl.AddOutput "Surface2", oObjRevolution
    
    If (m_bAngle = True) Then
        Dim oCircle As IngrGeom3D.Circle3d
        
        'Create the top plane
        Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                        dCentreX, dCentreY, dTopPlaneZ, _
                                                                        oUpVector.x, oUpVector.y, oUpVector.z, _
                                                                        dOuterRadius)
        Set oComplexString = New ComplexString3d
        oComplexString.AddCurve oCircle, True
        
        Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
        Set oComplexString = Nothing
        Set oComplexString = New ComplexString3d
        Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                        dCentreX, dCentreY, dTopPlaneZ, _
                                                                        oDownVector.x, oDownVector.y, oDownVector.z, _
                                                                        dInnerRadius)
        oComplexString.AddCurve oCircle, True
        oPlane.AddBoundary oComplexString
        m_outputColl.AddOutput "Surface3", oPlane
        Set oCircle = Nothing
        Set oComplexString = Nothing
        Set oPlane = Nothing
        
        'Create the bottom plane
        Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                        dCentreX, dCentreY, dBottomPlaneZ, _
                                                                        oDownVector.x, oDownVector.y, oDownVector.z, _
                                                                        dOuterRadius)
        Set oComplexString = New ComplexString3d
        oComplexString.AddCurve oCircle, True
        'Create a plane for the outer circle on top
        Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
        Set oComplexString = Nothing
        Set oComplexString = New ComplexString3d
        Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                        dCentreX, dCentreY, dBottomPlaneZ, _
                                                                        oUpVector.x, oUpVector.y, oUpVector.z, _
                                                                        dInnerRadius)
        oComplexString.AddCurve oCircle, True
        oPlane.AddBoundary oComplexString
        m_outputColl.AddOutput "Surface4", oPlane
        Set oCircle = Nothing
        Set oComplexString = Nothing
        Set oPlane = Nothing
        
    Else
        Dim oArc            As IngrGeom3D.Arc3d
        
        Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
        m_outputColl.AddOutput "Point1", oPoint3d
        
        Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, (oPoint3.x + oPoint4.x) / 2, (oPoint3.y + oPoint4.y) / 2, 0)
        m_outputColl.AddOutput "Point2", oPoint3d
        
        'Creation of surface 3
        Set oComplexString = New ComplexString3d
        oCenterPoint.z = dTopPlaneZ
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                oUpVector.x, oUpVector.y, oUpVector.z, _
                                oPoint8.x, oPoint8.y, oPoint8.z, _
                                oPoint5.x, oPoint5.y, oPoint5.z)
        oComplexString.AddCurve oArc, True
        Set oArc = Nothing
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oPoint7.x, oPoint7.y, oPoint7.z, _
                                     oPoint8.x, oPoint8.y, oPoint8.z)
        oComplexString.AddCurve oLine, True
        Set oLine = Nothing
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                oDownVector.x, oDownVector.y, oDownVector.z, _
                                oPoint6.x, oPoint6.y, oPoint6.z, _
                                oPoint7.x, oPoint7.y, oPoint7.z)
        oComplexString.AddCurve oArc, True
        Set oArc = Nothing
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oPoint5.x, oPoint5.y, oPoint5.z, _
                                    oPoint6.x, oPoint6.y, oPoint6.z)
        oComplexString.AddCurve oLine, True
        Set oLine = Nothing
    
        Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
        m_outputColl.AddOutput "Surface3", oPlane
        Set oComplexString = Nothing
                
        'Creation of Surface4
        
        Set oComplexString = New ComplexString3d
        oCenterPoint.z = dBottomPlaneZ

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oPoint1.x, oPoint1.y, oPoint1.z, _
                                    oPoint2.x, oPoint2.y, oPoint2.z)
        oComplexString.AddCurve oLine, True
        Set oLine = Nothing
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                    oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                    oDownVector.x, oDownVector.y, oDownVector.z, _
                                    oPoint2.x, oPoint2.y, oPoint2.z, _
                                    oPoint3.x, oPoint3.y, oPoint3.z)
        oComplexString.AddCurve oArc, True
        Set oArc = Nothing
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oPoint3.x, oPoint3.y, oPoint3.z, _
                                     oPoint4.x, oPoint4.y, oPoint4.z)
        oComplexString.AddCurve oLine, True
        Set oLine = Nothing
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                    oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                    oUpVector.x, oUpVector.y, oUpVector.z, _
                                    oPoint4.x, oPoint4.y, oPoint4.z, _
                                    oPoint1.x, oPoint1.y, oPoint1.z)
        oComplexString.AddCurve oArc, True
        Set oArc = Nothing
        
        Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oComplexString)
        
        m_outputColl.AddOutput "Surface4", oPlane
        
        arrPoints(0) = oPoint2.x
        arrPoints(1) = oPoint2.y
        arrPoints(2) = oPoint2.z
        arrPoints(3) = oPoint1.x
        arrPoints(4) = oPoint1.y
        arrPoints(5) = oPoint1.z
        arrPoints(6) = oPoint5.x
        arrPoints(7) = oPoint5.y
        arrPoints(8) = oPoint5.z
        arrPoints(9) = oPoint6.x
        arrPoints(10) = oPoint6.y
        arrPoints(11) = oPoint6.z
        
        'Creation of Edge1 to 4
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
        m_outputColl.AddOutput "Edge1", line
        Set line = Nothing
        
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
        m_outputColl.AddOutput "Edge2", line
        Set line = Nothing
        
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
        m_outputColl.AddOutput "Edge3", line
        Set line = Nothing
        
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), arrPoints(0), arrPoints(1), arrPoints(2))
        m_outputColl.AddOutput "Edge4", line
        Set line = Nothing
        
        'Creation of Surface5
        Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
        m_outputColl.AddOutput "Surface5", oPlane
        
        arrPoints(0) = oPoint3.x
        arrPoints(1) = oPoint3.y
        arrPoints(2) = oPoint3.z
        arrPoints(3) = oPoint7.x
        arrPoints(4) = oPoint7.y
        arrPoints(5) = oPoint7.z
        arrPoints(6) = oPoint8.x
        arrPoints(7) = oPoint8.y
        arrPoints(8) = oPoint8.z
        arrPoints(9) = oPoint4.x
        arrPoints(10) = oPoint4.y
        arrPoints(11) = oPoint4.z
              
        'Creation of Edge5 to 8
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
        m_outputColl.AddOutput "Edge5", line
        Set line = Nothing
        
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
        m_outputColl.AddOutput "Edge6", line
        Set line = Nothing
        
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
        m_outputColl.AddOutput "Edge7", line
        Set line = Nothing
        
        Set line = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), arrPoints(0), arrPoints(1), arrPoints(2))
        m_outputColl.AddOutput "Edge8", line
        Set line = Nothing
        
        'Creation of Surface6
        Set oPlane = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, arrPoints)
        m_outputColl.AddOutput "Surface6", oPlane
        
    End If
        
    Set oPoint3d = Nothing
    Set oObjRevolution = Nothing
    Set oLine = Nothing
    Set oPlane = Nothing
    Set oComplexString = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

''' This function creates persistent/Transient Line based on
''' start and end points of the line
'''<{(Line begin)}>
Private Function PlaceLine(pResourceMgr As IUnknown, ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition) _
                            As IngrGeom3D.Line3d
    
    On Error GoTo ErrorHandler
        
    Dim oLine As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    ' Create Line object
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(pResourceMgr, _
                                startPoint.x, startPoint.y, startPoint.z, _
                                endPoint.x, endPoint.y, endPoint.z)
    
    
    Set PlaceLine = oLine
    Set oLine = Nothing
    Set oGeomFactory = Nothing

    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(Line end)}>

''' This function creates persistent revolution based on curve
''' axis of revolution and angle
'''<{(Revolution begin)}>
Private Function PlaceRevolution(ByVal oResourceManager As IUnknown, _
                                ByVal objCurve As Object, _
                                ByVal axisVector As IJDVector, _
                                ByVal oCenterPoint As IJDPosition, _
                                revAngle As Double, _
                                isCapped As Boolean)

' Construction of revolution
   On Error GoTo ErrorHandler
        
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Dim objRevolution   As IngrGeom3D.Revolution3d
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objRevolution = oGeomFactory.Revolutions3d.CreateByCurve( _
                                                    oResourceManager, _
                                                    objCurve, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                    revAngle, isCapped)

    Set PlaceRevolution = objRevolution
    Set objRevolution = Nothing
    Set oGeomFactory = Nothing
    
    Exit Function

ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(Revolution end)}>

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

