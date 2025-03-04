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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         VRK
'   Creation Date:  Friday,March 9 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Eleven Outputs:
'      In which Default Surface is created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
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
    
    'Define Vessel Head types
    Const EndType_2TO1 = 1
    Const EndType_CONE = 2
    Const EndType_DOME = 3
    Const EndType_FnD = 4
    Const EndType_FLAT = 5
    Const EndType_FLGD = 6
    Const EndType_HEMI = 7
    Const EndType_NONE = 8
    Const EndType_TORC = 9
    Const EndType_TORS = 10
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput  As Double
    Dim ObjVessel As Object
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parBootHorLocation As Double
    Dim parBootVertLocation As Double
    Dim parBootDiameter As Double
    Dim parBootAngularLocation As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parThirdSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parStiffenerRadius As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
  
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parBootHorLocation = arrayOfInputs(6)           'P4
    parBootVertLocation = arrayOfInputs(7)          'P5
    parBootDiameter = arrayOfInputs(8)              'P6
    parBootAngularLocation = arrayOfInputs(9)       'P7
    parVesselCenterHeight = arrayOfInputs(10)       'P8
    parFirstSupportLocation = arrayOfInputs(11)     'P9
    parSecondSupportLocation = arrayOfInputs(12)    'P10
    parThirdSupportLocation = arrayOfInputs(13)     'P11
    parSupportThickness = arrayOfInputs(14)         'P12
    parSupportLength = arrayOfInputs(15)            'P13
    parStiffenerRadius = arrayOfInputs(16)          'P14
    parEndHeadConeTopDiameter = arrayOfInputs(17)   'P51
    parEndHeadConeHeight = arrayOfInputs(18)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(19)     'P51 for TORC n TORS
    parEndHeadDomeradius = arrayOfInputs(20)        'P50
    parEndHeadFlangedThick1 = arrayOfInputs(21)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(22)     'P51
    parEndHeadSphericalRadius = arrayOfInputs(23)   'P50
    parInsulationThickness = arrayOfInputs(24)
    parXBoltHole = arrayOfInputs(25)
    parYBoltHole = arrayOfInputs(26)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
     
    'Insert your code for output DefaultSurface    ''Place Default surface at Origin
    'Create non-persistent circle to use for creating default surface ---
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                0, 0, 0, _
                                0, 0, -1, _
                                0.1 * parVesselDiameter / 2)

    'Create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, objCircle)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    
    'Insert your code for Vessel
    stPoint.Set 0, 0, 0
    enPoint.Set parVesselTantoTan, 0, 0
    Set ObjVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, False)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
                                                           
    'Insert your code for output (Boot)
    If CmpDblGreaterthan(parBootHorLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootVertLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootDiameter, LINEAR_TOLERANCE) Then

                Dim oAxisVect As New AutoMath.DVector
                Dim oRevCenPt As New AutoMath.DPosition
                Dim cenX As Double
                Dim cenY As Double
                Dim cenZ As Double
                Dim MajorX As Double
                Dim MajorY As Double
                Dim MajorZ As Double
                Dim mMRatio As Double
                Dim StartAngle As Double
                Dim SweepAngle As Double
                Dim normalX As Double
                Dim normalY As Double
                Dim normalZ As Double

                'Boot cylinder
                Dim ObjBootVessel As Object
                Dim meetDistance As Double
                meetDistance = Sqr((parVesselDiameter / 2) ^ 2 - (parBootDiameter / 2) ^ 2)
                stPoint.Set parBootHorLocation, meetDistance * Sin(parBootAngularLocation), _
                                            meetDistance * Cos(parBootAngularLocation)
                enPoint.Set parBootHorLocation, parBootVertLocation * Sin(parBootAngularLocation), _
                                            parBootVertLocation * Cos(parBootAngularLocation)
                Set ObjBootVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBootDiameter, True)
                'Set the output
                iOutput = iOutput + 1
                m_OutputColl.AddOutput "BootVessel", ObjBootVessel
                Set ObjBootVessel = Nothing

                'Boot Dome
                Dim ObjBootDome As IngrGeom3D.Revolution3d
                
                cenX = parBootHorLocation
                cenY = 0
                cenZ = parBootVertLocation

                MajorX = 0
                MajorY = (parBootDiameter / 2)
                MajorZ = 0

                mMRatio = 0.5
                StartAngle = PI / 2   '-0.5 * PI
                SweepAngle = PI / 2

                normalX = 1
                normalY = 0
                normalZ = 0

                Dim oBootBottomArc As IngrGeom3D.EllipticalArc3d
                Set oBootBottomArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                    cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                                    StartAngle, SweepAngle)

                oAxisVect.Set 0, 0, 1

                Dim BootRevCenPt As New AutoMath.DPosition
                BootRevCenPt.Set parBootHorLocation, 0, parBootVertLocation
                Set ObjBootDome = PlaceRevolution(m_OutputColl, oBootBottomArc, oAxisVect, BootRevCenPt, 2 * PI, True)

                'Apply rotation transformations to the object created.
                Dim oTransMat     As New AutoMath.DT4x4
                Dim oRotVec As New AutoMath.DVector
                oTransMat.LoadIdentity
                oRotVec.Set 1, 0, 0

                oTransMat.Rotate -parBootAngularLocation, oRotVec
                ObjBootDome.Transform oTransMat

                'Set the output
                iOutput = iOutput + 1
                m_OutputColl.AddOutput "BootDome", ObjBootDome
                Set ObjBootDome = Nothing

                Set BootRevCenPt = Nothing
                Set oBootBottomArc = Nothing
                Set oTransMat = Nothing
                Set oRotVec = Nothing
    End If

    'Placement of supports
    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) Then
                If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
                    'Insert your code for Support1
                    Dim ObjSupport1 As Object
                    stPoint.Set parFirstSupportLocation - parSupportThickness / 2, _
                                 -parSupportLength / 2, _
                                 -parVesselCenterHeight
                    enPoint.Set stPoint.x + parSupportThickness, _
                                stPoint.y + parSupportLength, _
                                stPoint.z + parVesselCenterHeight
                    Set ObjSupport1 = PlaceBox(m_OutputColl, stPoint, enPoint)
                    'Set the output
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Support_", ObjSupport1
                    Set ObjSupport1 = Nothing
                End If

                If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
                    'Insert your code for Support2
                    Dim ObjSupport2 As Object
                    stPoint.Set parFirstSupportLocation + parSecondSupportLocation - parSupportThickness / 2, _
                                 -parSupportLength / 2, _
                                 -parVesselCenterHeight
                    enPoint.Set stPoint.x + parSupportThickness, _
                                stPoint.y + parSupportLength, _
                                stPoint.z + parVesselCenterHeight
                    Set ObjSupport2 = PlaceBox(m_OutputColl, stPoint, enPoint)
                    'Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport2
                    Set ObjSupport2 = Nothing
                End If

                If CmpDblGreaterthan(parThirdSupportLocation, LINEAR_TOLERANCE) Then
                    'Insert your code for output Support3
                    Dim ObjSupport3 As Object
                    stPoint.Set parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation - parSupportThickness / 2, _
                                 -parSupportLength / 2, _
                                 -parVesselCenterHeight
                    enPoint.Set stPoint.x + parSupportThickness, _
                                stPoint.y + parSupportLength, _
                                stPoint.z + parVesselCenterHeight
                    Set ObjSupport3 = PlaceBox(m_OutputColl, stPoint, enPoint)

                    'Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport3
                    Set ObjSupport3 = Nothing
                End If
    End If

    'Placement of stiffeners
    If CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) And CmpDblGreaterthan(parStiffenerRadius, LINEAR_TOLERANCE) Then
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
            'Insert your code for Stiffner1
            Dim ObjStiffner1 As Object
            Dim StiffnerPoints(0 To 14)  As Double

            'Consider a section of Stiffener in XZ plane
            'Left hand side Bottom point
            StiffnerPoints(0) = parFirstSupportLocation - parSupportThickness / 2
            StiffnerPoints(1) = 0
            StiffnerPoints(2) = parVesselDiameter / 2

            'Right hand side Bottom point
            StiffnerPoints(3) = parFirstSupportLocation + parSupportThickness / 2
            StiffnerPoints(4) = 0
            StiffnerPoints(5) = parVesselDiameter / 2

            'Right hand side Top point
            StiffnerPoints(6) = parFirstSupportLocation + parSupportThickness / 2
            StiffnerPoints(7) = 0
            StiffnerPoints(8) = parStiffenerRadius

            'Left hand side Top point
            StiffnerPoints(9) = parFirstSupportLocation - parSupportThickness / 2
            StiffnerPoints(10) = 0
            StiffnerPoints(11) = parStiffenerRadius

            'Left hand side Bottom point
            StiffnerPoints(12) = StiffnerPoints(0)
            StiffnerPoints(13) = StiffnerPoints(1)
            StiffnerPoints(14) = StiffnerPoints(2)

            'Prepare profile
            oAxisVect.Set 1, 0, 0
            oRevCenPt.Set parFirstSupportLocation, 0, 0
            Dim oLineString As IngrGeom3D.LineString3d
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
            Set ObjStiffner1 = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oRevCenPt, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Stiffner_", ObjStiffner1
            Set ObjStiffner1 = Nothing
            Set oLineString = Nothing
        End If

        If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
            'Insert your code for Stiffner2
            Dim ObjStiffner2 As Object
            'Left hand side Bottom point
            StiffnerPoints(0) = parFirstSupportLocation + parSecondSupportLocation - _
                                    parSupportThickness / 2
            StiffnerPoints(1) = 0
            StiffnerPoints(2) = parVesselDiameter / 2

            'Right hand side Bottom point
            StiffnerPoints(3) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parSupportThickness / 2
            StiffnerPoints(4) = 0
            StiffnerPoints(5) = parVesselDiameter / 2

            'Right hand side Top point
            StiffnerPoints(6) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parSupportThickness / 2
            StiffnerPoints(7) = 0
            StiffnerPoints(8) = parStiffenerRadius

            'Left hand side Top point
            StiffnerPoints(9) = parFirstSupportLocation + parSecondSupportLocation - _
                                    parSupportThickness / 2
            StiffnerPoints(10) = 0
            StiffnerPoints(11) = parStiffenerRadius

            'Left hand side Bottom point
            StiffnerPoints(12) = StiffnerPoints(0)
            StiffnerPoints(13) = StiffnerPoints(1)
            StiffnerPoints(14) = StiffnerPoints(2)

            Set oLineString = New IngrGeom3D.LineString3d
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
            oAxisVect.Set 1, 0, 0
            oRevCenPt.Set parFirstSupportLocation + parSecondSupportLocation, 0, 0
            Set ObjStiffner2 = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oRevCenPt, 2 * PI, True)

            'Set the output
            m_OutputColl.AddOutput "Stiffner_", ObjStiffner2
            Set ObjStiffner2 = Nothing

            'Remove  linestring
            Dim objLineString As IJDObject
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
    End If

    If CmpDblGreaterthan(parThirdSupportLocation, LINEAR_TOLERANCE) Then
            'Insert your code for Stiffner3
            Dim ObjStiffner3 As Object
            'Left hand side Bottom point
            StiffnerPoints(0) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation - parSupportThickness / 2
            StiffnerPoints(1) = 0
            StiffnerPoints(2) = parVesselDiameter / 2

            'Right hand side Bottom point
            StiffnerPoints(3) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation + parSupportThickness / 2
            StiffnerPoints(4) = 0
            StiffnerPoints(5) = parVesselDiameter / 2

            'Right hand side Top point
            StiffnerPoints(6) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation + parSupportThickness / 2
            StiffnerPoints(7) = 0
            StiffnerPoints(8) = parStiffenerRadius

            'Left hand side Top point
            StiffnerPoints(9) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation - parSupportThickness / 2
            StiffnerPoints(10) = 0
            StiffnerPoints(11) = parStiffenerRadius

            'Left hand side Bottom point
            StiffnerPoints(12) = StiffnerPoints(0)
            StiffnerPoints(13) = StiffnerPoints(1)
            StiffnerPoints(14) = StiffnerPoints(2)

            Set oLineString = New IngrGeom3D.LineString3d
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
            oAxisVect.Set 1, 0, 0
            oRevCenPt.Set parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation, 0, 0
            Set ObjStiffner3 = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oRevCenPt, 2 * PI, True)

            'Set the output
            m_OutputColl.AddOutput "Stiffner_", ObjStiffner3
            Set ObjStiffner3 = Nothing

            'Remove  linestring
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
            Set objLineString = Nothing
        End If
    End If
    Set oAxisVect = Nothing
    Set oRevCenPt = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    'Inser your code for creating the edges for Support(s)
    Dim iCount As Integer
    Dim jCount As Integer
    Dim SupportPoints(0 To 23) As Double
    Dim ObjEdgeColl As Collection
    Set ObjEdgeColl = New Collection
    
    Dim SupportCenter As New AutoMath.DPosition
    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) Then
        SupportCenter.Set parFirstSupportLocation, 0, -parVesselCenterHeight
                For iCount = 0 To 12 Step 12
                    SupportPoints(0 + iCount) = SupportCenter.x - parSupportThickness / 2
                    SupportPoints(1 + iCount) = SupportCenter.y - parSupportLength / 2
                    SupportPoints(2 + iCount) = SupportCenter.z
            
                    SupportPoints(3 + iCount) = SupportCenter.x - parSupportThickness / 2
                    SupportPoints(4 + iCount) = SupportCenter.y + parSupportLength / 2
                    SupportPoints(5 + iCount) = SupportCenter.z
                    
                    SupportPoints(6 + iCount) = SupportCenter.x + parSupportThickness / 2
                    SupportPoints(7 + iCount) = SupportCenter.y + parSupportLength / 2
                    SupportPoints(8 + iCount) = SupportCenter.z
                    
                    SupportPoints(9 + iCount) = SupportCenter.x + parSupportThickness / 2
                    SupportPoints(10 + iCount) = SupportCenter.y - parSupportLength / 2
                    SupportPoints(11 + iCount) = SupportCenter.z
        
                    SupportCenter.z = SupportCenter.z + parVesselCenterHeight
                Next iCount
               
        'Edges
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then

                 'Edge1 Point1 to Point2
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5))
                 'Edge2 Point1 to Point4
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11))
                 'Edge3 Point1 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge4 Point4 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge5 Point2 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge6 Point4 to Point8
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11), _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23))
                 'Edge7 Point6 to Point2
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5))
                 'Edge8 Point6 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge9 Point6 to Point7
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20))
                 'Edge10 Point7 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge11 Point8 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge12 Point7 to Point8
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20), _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23))
                     
                'Set the Output
                iOutput = iOutput + 1
                For iCount = 1 To ObjEdgeColl.Count
                    m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
                Next iCount
                Set ObjEdgeColl = Nothing
        End If
       
        'Support2
        SupportCenter.Set parFirstSupportLocation + parSecondSupportLocation, 0, -parVesselCenterHeight
                For iCount = 0 To 12 Step 12
                    SupportPoints(0 + iCount) = SupportCenter.x - parSupportThickness / 2
                    SupportPoints(1 + iCount) = SupportCenter.y - parSupportLength / 2
                    SupportPoints(2 + iCount) = SupportCenter.z
            
                    SupportPoints(3 + iCount) = SupportCenter.x - parSupportThickness / 2
                    SupportPoints(4 + iCount) = SupportCenter.y + parSupportLength / 2
                    SupportPoints(5 + iCount) = SupportCenter.z
                    
                    SupportPoints(6 + iCount) = SupportCenter.x + parSupportThickness / 2
                    SupportPoints(7 + iCount) = SupportCenter.y + parSupportLength / 2
                    SupportPoints(8 + iCount) = SupportCenter.z
                    
                    SupportPoints(9 + iCount) = SupportCenter.x + parSupportThickness / 2
                    SupportPoints(10 + iCount) = SupportCenter.y - parSupportLength / 2
                    SupportPoints(11 + iCount) = SupportCenter.z
        
                    SupportCenter.z = SupportCenter.z + parVesselCenterHeight
                Next iCount
               
        'Edges
        If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
                 Set ObjEdgeColl = New Collection

                 'Edge1 Point1 to Point2
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5))
                 'Edge2 Point1 to Point4
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11))
                 'Edge3 Point1 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge4 Point4 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge5 Point2 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge6 Point4 to Point8
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11), _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23))
                 'Edge7 Point6 to Point2
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5))
                 'Edge8 Point6 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge9 Point6 to Point7
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20))
                 'Edge10 Point7 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge11 Point8 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge12 Point7 to Point8
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20), _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23))
                     
                'Set the Output
                For iCount = 1 To ObjEdgeColl.Count
                    m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
                Next iCount
                Set ObjEdgeColl = Nothing
        End If
        
        'Support3
        SupportCenter.Set parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation, 0, -parVesselCenterHeight
                For iCount = 0 To 12 Step 12
                    SupportPoints(0 + iCount) = SupportCenter.x - parSupportThickness / 2
                    SupportPoints(1 + iCount) = SupportCenter.y - parSupportLength / 2
                    SupportPoints(2 + iCount) = SupportCenter.z
            
                    SupportPoints(3 + iCount) = SupportCenter.x - parSupportThickness / 2
                    SupportPoints(4 + iCount) = SupportCenter.y + parSupportLength / 2
                    SupportPoints(5 + iCount) = SupportCenter.z
                    
                    SupportPoints(6 + iCount) = SupportCenter.x + parSupportThickness / 2
                    SupportPoints(7 + iCount) = SupportCenter.y + parSupportLength / 2
                    SupportPoints(8 + iCount) = SupportCenter.z
                    
                    SupportPoints(9 + iCount) = SupportCenter.x + parSupportThickness / 2
                    SupportPoints(10 + iCount) = SupportCenter.y - parSupportLength / 2
                    SupportPoints(11 + iCount) = SupportCenter.z
        
                    SupportCenter.z = SupportCenter.z + parVesselCenterHeight
                Next iCount
               
        'Edges
        If CmpDblGreaterthan(parThirdSupportLocation, LINEAR_TOLERANCE) Then
                 Set ObjEdgeColl = New Collection

                 'Edge1 Point1 to Point2
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5))
                 'Edge2 Point1 to Point4
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11))
                 'Edge3 Point1 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(0), SupportPoints(1), SupportPoints(2), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge4 Point4 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge5 Point2 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge6 Point4 to Point8
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(9), SupportPoints(10), SupportPoints(11), _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23))
                 'Edge7 Point6 to Point2
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(3), SupportPoints(4), SupportPoints(5))
                 'Edge8 Point6 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge9 Point6 to Point7
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(15), SupportPoints(16), SupportPoints(17), _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20))
                 'Edge10 Point7 to Point3
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20), _
                                             SupportPoints(6), SupportPoints(7), SupportPoints(8))
                 'Edge11 Point8 to Point5
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23), _
                                             SupportPoints(12), SupportPoints(13), SupportPoints(14))
                 'Edge12 Point7 to Point8
                 ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                             SupportPoints(18), SupportPoints(19), SupportPoints(20), _
                                             SupportPoints(21), SupportPoints(22), SupportPoints(23))
                     
                'Set the Output
                For iCount = 1 To ObjEdgeColl.Count
                    m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
                Next iCount
                Set ObjEdgeColl = Nothing
        End If
        
        
        'Creating the Points on the SupportCenter of each Surface of Supports
        Dim ObjPointColl As Collection
        Set ObjPointColl = New Collection
    
         If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
             SupportCenter.Set parFirstSupportLocation, 0, -parVesselCenterHeight

            'Points on Right and Left surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x + parSupportThickness / 2, SupportCenter.y, SupportCenter.z / 2)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x - parSupportThickness / 2, SupportCenter.y, SupportCenter.z / 2)
            'Points on Front and Back surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y - parSupportLength / 2, SupportCenter.z / 2)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y + parSupportLength / 2, SupportCenter.z / 2)
            'Points on Top and Bottom surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y, SupportCenter.z + parVesselCenterHeight)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y, SupportCenter.z)
            'Set the output
            iOutput = iOutput + 1
            For iCount = 1 To ObjPointColl.Count
                m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
            Next iCount
            Set ObjPointColl = Nothing
        End If
       'Creating the Points on the SupportCenter of each Surface of Supports
        Set ObjPointColl = New Collection
         If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
             SupportCenter.Set parFirstSupportLocation + parSecondSupportLocation, 0, -parVesselCenterHeight

            'Points on Right and Left surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x + parSupportThickness / 2, SupportCenter.y, SupportCenter.z / 2)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x - parSupportThickness / 2, SupportCenter.y, SupportCenter.z / 2)
            'Points on Front and Back surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y - parSupportLength / 2, SupportCenter.z / 2)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y + parSupportLength / 2, SupportCenter.z / 2)
            'Points on Top and Bottom surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y, SupportCenter.z + parVesselCenterHeight)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y, SupportCenter.z)
            'Set the output
            For iCount = 1 To ObjPointColl.Count
                m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
            Next iCount
            Set ObjPointColl = Nothing
        End If
        
        'Creating the Points on the SupportCenter of each Surface of Supports
        Set ObjPointColl = New Collection
         If CmpDblGreaterthan(parThirdSupportLocation, LINEAR_TOLERANCE) Then
             SupportCenter.Set parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation, 0, -parVesselCenterHeight

            'Points on Right and Left surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x + parSupportThickness / 2, SupportCenter.y, SupportCenter.z / 2)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x - parSupportThickness / 2, SupportCenter.y, SupportCenter.z / 2)
            'Points on Front and Back surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y - parSupportLength / 2, SupportCenter.z / 2)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y + parSupportLength / 2, SupportCenter.z / 2)
            'Points on Top and Bottom surfaces
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y, SupportCenter.z + parVesselCenterHeight)
            ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    SupportCenter.x, SupportCenter.y, SupportCenter.z)
            'Set the output
            For iCount = 1 To ObjPointColl.Count
                m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
            Next iCount
            Set ObjPointColl = Nothing
        End If
        Set oGeomFactory = Nothing
    End If

    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Foundation Port1
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
   'The origin of the port is taken to be at the corner point of the support base.
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
    Dim holes() As Variant
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
    
    'Foundation Port2
    'The origin of the port is taken to be at the corner point of the support base.
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

    'Foundation Port3
    'The origin of the port is taken to be at the corner point of the support base.
    dOrigin(0) = parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation
    dOrigin(1) = 0
    dOrigin(2) = -parVesselCenterHeight

    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort3", _
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

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


