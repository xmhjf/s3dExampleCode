VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Wednesday, March 7 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Eight Outputs:
'      In which Default Surface is created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjVessel       As Object
    Dim ObjVesselLEH    As Object
    Dim ObjVesselREH    As Object
    Dim iOutput         As Double
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double


    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)        'P2
    parVesselDiameter = arrayOfInputs(5)        'P3
    parVesselCenterHeight = arrayOfInputs(6)    'P4
    parFirstSupportLocation = arrayOfInputs(7)  'P5
    parSecondSupportLocation = arrayOfInputs(8) 'P6
    parSupportThickness = arrayOfInputs(9)      'P7
    parSupportLength = arrayOfInputs(10)        'P8
    parXBoltHole = arrayOfInputs(11)
    parYBoltHole = arrayOfInputs(12)
    parEndHeadConeTopDiameter = arrayOfInputs(13)
    parEndHeadConeHeight = arrayOfInputs(14)
    parEndHeadKnuckleRadius = arrayOfInputs(15)
    parEndHeadDomeradius = arrayOfInputs(16)
    parEndHeadFlangedThick1 = arrayOfInputs(17)
    parEndHeadFlangedThick2 = arrayOfInputs(18)
    parEndHeadSphericalRadius = arrayOfInputs(19)
    parInsulationThickness = arrayOfInputs(20)

    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'Insert your code for DefaultSurface
    'Create non-persistent circle to use for creating default surface ---
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim objPlane    As IngrGeom3D.Plane3d
    Dim objCircle   As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, -1, _
                            0.1 * parVesselDiameter / 2)

    'Create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objCircle = Nothing

    'Origin is taken at left dome center point (i.e. Placement point 1)
    'Vessel is formed by creating a Cylinder and two Elliptical Heads rotating about X-Axis

    'Insert your code for Vessel
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set parVesselTantoTan, 0, 0
    Set ObjVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    'Insert your code for Supports
    Dim NormalVect As New AutoMath.DVector
    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        Dim ObjSupport As Object
        Dim ObjSurface As Object
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
            'Place First Support
            Dim lineStrPoints(0 To 14)  As Double
            lineStrPoints(0) = parFirstSupportLocation - parSupportThickness / 2
            lineStrPoints(1) = -parSupportLength / 2
            lineStrPoints(2) = -parVesselCenterHeight
            
            lineStrPoints(3) = lineStrPoints(0)
            lineStrPoints(4) = lineStrPoints(1) + parSupportLength
            lineStrPoints(5) = lineStrPoints(2)

            lineStrPoints(6) = lineStrPoints(3) + parSupportThickness
            lineStrPoints(7) = lineStrPoints(4)
            lineStrPoints(8) = lineStrPoints(5)

            lineStrPoints(9) = lineStrPoints(6)
            lineStrPoints(10) = lineStrPoints(7) - parSupportLength
            lineStrPoints(11) = lineStrPoints(8)

            lineStrPoints(12) = lineStrPoints(0)
            lineStrPoints(13) = lineStrPoints(1)
            lineStrPoints(14) = lineStrPoints(2)
            
            NormalVect.Set 0, 0, 1
            Set ObjSupport = PlaceProjectedBox(m_OutputColl, lineStrPoints, NormalVect, parVesselCenterHeight, False)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Support_", ObjSupport
            Set ObjSupport = Nothing

            'Create First DefaultSurface
            Set ObjSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Surface_", ObjSurface
            Set ObjSurface = Nothing
        End If

        If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
            'Place Second Support
            lineStrPoints(0) = parFirstSupportLocation + parSecondSupportLocation _
                                - parSupportThickness / 2
            lineStrPoints(1) = -parSupportLength / 2
            lineStrPoints(2) = -parVesselCenterHeight
            
            lineStrPoints(3) = lineStrPoints(0)
            lineStrPoints(4) = lineStrPoints(1) + parSupportLength
            lineStrPoints(5) = lineStrPoints(2)

            lineStrPoints(6) = lineStrPoints(3) + parSupportThickness
            lineStrPoints(7) = lineStrPoints(4)
            lineStrPoints(8) = lineStrPoints(5)

            lineStrPoints(9) = lineStrPoints(6)
            lineStrPoints(10) = lineStrPoints(7) - parSupportLength
            lineStrPoints(11) = lineStrPoints(8)
            
            lineStrPoints(12) = lineStrPoints(0)
            lineStrPoints(13) = lineStrPoints(1)
            lineStrPoints(14) = lineStrPoints(2)
            
            Set ObjSupport = PlaceProjectedBox(m_OutputColl, lineStrPoints, NormalVect, parVesselCenterHeight, False)

            'Set the output
            m_OutputColl.AddOutput "Support_", ObjSupport
            Set ObjSupport = Nothing

            'Create Second DefaultSurface
            Set ObjSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

            'Set the output
            m_OutputColl.AddOutput "Surface_", ObjSurface
            Set ObjSurface = Nothing
        End If
    End If
    
    'Inser your code for creating the Edges and Points for Supports
    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Or CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
            Dim iCount As Integer
            Dim jCount As Integer
            Dim Points(0 To 1, 0 To 23) As Double
            Dim SupportCenter As New AutoMath.DPosition
            
            'Locating Points at Corners of each Support
            SupportCenter.Set parFirstSupportLocation - parSupportThickness / 2, 0, -parVesselCenterHeight / 2
            jCount = 0
            While jCount <= 1
                If CmpDblLessThan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
                    SupportCenter.x = parFirstSupportLocation + parSecondSupportLocation - parSupportThickness / 2
                    jCount = 1
                End If
                For iCount = 0 To 12 Step 12
                    Points(jCount, 0 + iCount) = SupportCenter.x
                    Points(jCount, 1 + iCount) = SupportCenter.y + parSupportLength / 2
                    Points(jCount, 2 + iCount) = SupportCenter.z + parVesselCenterHeight / 2
            
                    Points(jCount, 3 + iCount) = SupportCenter.x
                    Points(jCount, 4 + iCount) = SupportCenter.y + parSupportLength / 2
                    Points(jCount, 5 + iCount) = SupportCenter.z - parVesselCenterHeight / 2
                    
                    Points(jCount, 6 + iCount) = SupportCenter.x
                    Points(jCount, 7 + iCount) = SupportCenter.y - parSupportLength / 2
                    Points(jCount, 8 + iCount) = SupportCenter.z + parVesselCenterHeight / 2
                    
                    Points(jCount, 9 + iCount) = SupportCenter.x
                    Points(jCount, 10 + iCount) = SupportCenter.y - parSupportLength / 2
                    Points(jCount, 11 + iCount) = SupportCenter.z - parVesselCenterHeight / 2
                    
                    SupportCenter.x = SupportCenter.x + parSupportThickness
                Next iCount
                If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) And jCount = 0 Then
                    SupportCenter.x = parFirstSupportLocation + parSecondSupportLocation - parSupportThickness / 2
                Else
                    jCount = jCount + 1
                End If
                jCount = jCount + 1
            Wend
            
            'Edges
            Dim ObjEdgeColl As New Collection
            'Edge1 Point1 to Point2
            jCount = 0
            While jCount <= 1
                If CmpDblLessThan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
                    jCount = 1
                End If
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 0), Points(jCount, 1), Points(jCount, 2), _
                                            Points(jCount, 3), Points(jCount, 4), Points(jCount, 5))
                'Edge2 Point1 to Point3
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 0), Points(jCount, 1), Points(jCount, 2), _
                                            Points(jCount, 6), Points(jCount, 7), Points(jCount, 8))
                'Edge3 Point1 to Point5
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 0), Points(jCount, 1), Points(jCount, 2), _
                                            Points(jCount, 12), Points(jCount, 13), Points(jCount, 14))
                'Edge4 Point4 to Point2
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 9), Points(jCount, 10), Points(jCount, 11), _
                                            Points(jCount, 3), Points(jCount, 4), Points(jCount, 5))
                'Edge5 Point4 to Point3
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 9), Points(jCount, 10), Points(jCount, 11), _
                                            Points(jCount, 6), Points(jCount, 7), Points(jCount, 8))
                'Edge6 Point4 to Point8
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 9), Points(jCount, 10), Points(jCount, 11), _
                                            Points(jCount, 21), Points(jCount, 22), Points(jCount, 23))
                'Edge7 Point6 to Point2
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 15), Points(jCount, 16), Points(jCount, 17), _
                                            Points(jCount, 3), Points(jCount, 4), Points(jCount, 5))
                'Edge8 Point6 to Point5
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 15), Points(jCount, 16), Points(jCount, 17), _
                                            Points(jCount, 12), Points(jCount, 13), Points(jCount, 14))
                'Edge9 Point6 to Point8
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 15), Points(jCount, 16), Points(jCount, 17), _
                                            Points(jCount, 21), Points(jCount, 22), Points(jCount, 23))
                'Edge10 Point7 to Point3
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 18), Points(jCount, 19), Points(jCount, 20), _
                                            Points(jCount, 6), Points(jCount, 7), Points(jCount, 8))
                'Edge11 Point7 to Point5
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 18), Points(jCount, 19), Points(jCount, 20), _
                                            Points(jCount, 12), Points(jCount, 13), Points(jCount, 14))
                'Edge12 Point7 to Point8
                ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                            Points(jCount, 18), Points(jCount, 19), Points(jCount, 20), _
                                            Points(jCount, 21), Points(jCount, 22), Points(jCount, 23))
                If CmpDblLessThan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
                    jCount = jCount + 1
                End If
                jCount = jCount + 1
            Wend
            
            'Set the Output
            iOutput = iOutput + 1
            For iCount = 1 To ObjEdgeColl.Count
                m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
            Next iCount
            Set ObjEdgeColl = Nothing
            
            'Creating the Points on the Center of each Surface of Supports
            Dim ObjPointColl As Collection
            Set ObjPointColl = New Collection
            
            SupportCenter.Set parFirstSupportLocation, 0, -parVesselCenterHeight / 2
            iCount = 0
            While iCount <= 1
                If CmpDblLessThan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
                    SupportCenter.Set parFirstSupportLocation + parSecondSupportLocation, 0, -parVesselCenterHeight / 2
                    iCount = iCount + 1
                End If
                'Points on Right and Left surfaces
                ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                        SupportCenter.x + parSupportThickness / 2, SupportCenter.y, SupportCenter.z)
                ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                        SupportCenter.x - parSupportThickness / 2, SupportCenter.y, SupportCenter.z)
                'Points on Front and Back surfaces
                ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                        SupportCenter.x, SupportCenter.y - parSupportLength / 2, SupportCenter.z)
                ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                        SupportCenter.x, SupportCenter.y + parSupportLength / 2, SupportCenter.z)
                'Points on Top and Bottom surfaces
                ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                        SupportCenter.x, SupportCenter.y, SupportCenter.z + parVesselCenterHeight / 2)
                ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                        SupportCenter.x, SupportCenter.y, SupportCenter.z - parVesselCenterHeight / 2)
                If CmpDblLessThan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
                    iCount = iCount + 1
                End If
                SupportCenter.Set parFirstSupportLocation + parSecondSupportLocation, 0, -parVesselCenterHeight / 2
                iCount = iCount + 1
            Wend
            
            'Set the output
            iOutput = iOutput + 1
            For iCount = 1 To ObjPointColl.Count
                m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
            Next iCount
            Set ObjPointColl = Nothing
            Set SupportCenter = Nothing
        End If
    End If
    
    'Create Equipment Foundation Ports  at the bottom of Supports
    'Coordinate system is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |   X   |
    '   |                  |   ^   |
    '   |                  |   |   |
    '   -----> X           |   |   |Port CS
    '   Symbol CS          |   |----> Y
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |-------|
    
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    Dim holes() As Variant

    'Create FoundationPort1 at the bottom of Support1
    'The origin of the port is taken to be at the center point of the support base.
    dOrigin(0) = parFirstSupportLocation
    dOrigin(1) = 0
    dOrigin(2) = -parVesselCenterHeight
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Call ObjFoundationPort.GetHoles(holes())
    
    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole
    
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                             dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    
    'Create FoundationPort2 at the bottom Support2
    'The origin of the port is taken to be at the center point of the support base.
    dOrigin(0) = parFirstSupportLocation + parSecondSupportLocation
    dOrigin(1) = 0
    dOrigin(2) = -parVesselCenterHeight
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort2", _
                                                    False, m_OutputColl.ResourceManager)
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole
    
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                        dPoints() As Double, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean) As Object
    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler
      
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d
        
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = oGeomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLineStr, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    height, isCapped)
    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set oGeomFactory = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function
