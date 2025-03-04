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
'   Author:    VRK
'   Creation Date:  FriDay,March 23 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has five Outputs:
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
    Dim iOutput As Double
    Dim ObjVesselCylBody As Object
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportReferencePoint  As Double
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
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parSupportAngularLocation = arrayOfInputs(6)    'P4
    parNumberOfSupports = arrayOfInputs(7)          'P5
    parSupportHeight = arrayOfInputs(8)             'P6
    parSupportLength = arrayOfInputs(9)             'P7
    parSupportThickness = arrayOfInputs(10)         'P8
    parSupportReferencePoint = arrayOfInputs(11)    'P9
    parEndHeadConeTopDiameter = arrayOfInputs(12)   'P51
    parEndHeadConeHeight = arrayOfInputs(13)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(14)     'P51 for TORC and TORS
    parEndHeadDomeradius = arrayOfInputs(15)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(16)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(17)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(18)   'P50
    parInsulationThickness = arrayOfInputs(19)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory  As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for (Default Surface /plane at Control Point of vessel )
    'Create non-persistent circle to use for creating default surface ---
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                0, 0, 0, _
                                0, 0, -1, _
                                0.1 * parVesselDiameter / 2)

    'Create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing

    'Origin is at Place Point 1.
    'Insert your code for Vessel Cylindrical Body
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, 0, parVesselTantoTan
    Set ObjVesselCylBody = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parVesselDiameter, False)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselCylBody
    Set ObjVesselCylBody = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Place Circular FoundationPort
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = -parSupportReferencePoint
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Insert your code for output Leg Support
    Dim IncludedAng As Double
    Dim dYaxisToSupportAngle As Double
    Dim iCount As Integer
    Dim jCount As Integer
    If (parNumberOfSupports > 0) And CmpDblGreaterthan(parSupportHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) _
                                                    And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) Then
            
            IncludedAng = (PI * 2) / parNumberOfSupports
            dYaxisToSupportAngle = parSupportAngularLocation
            
            Dim topSurPts(0 To 3) As IJDPosition
            Dim botSurPts(0 To 3) As IJDPosition
            
            For iCount = 0 To 3
                Set topSurPts(iCount) = New DPosition
                Set botSurPts(iCount) = New DPosition
            Next iCount
            
            For iCount = 0 To parNumberOfSupports - 1
                    'form Line String and Project it
                    'Point 1
                    botSurPts(0).x = Sin(dYaxisToSupportAngle) * (parVesselDiameter / 2) + Cos(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(0).y = Cos(dYaxisToSupportAngle) * (parVesselDiameter / 2) - Sin(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(0).z = -parSupportReferencePoint
                    
                    'Point 2
                    botSurPts(1).x = botSurPts(0).x + Sin(dYaxisToSupportAngle) * parSupportThickness
                    botSurPts(1).y = botSurPts(0).y + parSupportThickness * Cos(dYaxisToSupportAngle)
                    botSurPts(1).z = botSurPts(0).z
                    
                    'Point 4
                    botSurPts(3).x = Sin(dYaxisToSupportAngle) * (parVesselDiameter / 2) - Cos(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(3).y = Cos(dYaxisToSupportAngle) * (parVesselDiameter / 2) + Sin(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(3).z = botSurPts(0).z
                    
                    'Point3
                    botSurPts(2).x = botSurPts(3).x + parSupportThickness * Sin(dYaxisToSupportAngle)
                    botSurPts(2).y = botSurPts(3).y + parSupportThickness * Cos(dYaxisToSupportAngle)
                    botSurPts(2).z = botSurPts(0).z
                    
                    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parSupportHeight
                    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, topSurPts(0).z
                    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, topSurPts(0).z
                    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, topSurPts(0).z
                    
                    Dim ObjCollection As Collection
                    Set ObjCollection = PlaceTrapezoidWithPlanes(m_OutputColl, topSurPts, botSurPts)
                    
                    iOutput = iOutput + 1
                    For jCount = 1 To ObjCollection.Count
                            m_OutputColl.AddOutput "Planes_", ObjCollection(jCount)
                    Next jCount
    
                    Set ObjCollection = Nothing
                   
                    'Increment the Angle
                    dYaxisToSupportAngle = dYaxisToSupportAngle + IncludedAng
                    
                    'Create the Edges(Lines) by two points at Bottom Surface
                        Dim ObjEdge As Object
                        Dim kCount As Integer
                        For kCount = 0 To 3
                            If kCount = 3 Then
                                Set ObjEdge = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, botSurPts(kCount).x, botSurPts(kCount).y, botSurPts(kCount).z, botSurPts(0).x, botSurPts(0).y, botSurPts(0).z)
                            Else
                                Set ObjEdge = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, botSurPts(kCount).x, botSurPts(kCount).y, botSurPts(kCount).z, botSurPts(kCount + 1).x, botSurPts(kCount + 1).y, botSurPts(kCount + 1).z)
                            End If
                        'Set the Output for Edges
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Edge_", ObjEdge
                        Set ObjEdge = Nothing
                        Next kCount
                       
                        Dim ObjPoint As Object
                        Dim x, y, z As Double
                        'Left FacePoint
                        x = (botSurPts(0).x + topSurPts(3).x) / 2
                        y = (botSurPts(0).y + topSurPts(1).y) / 2
                        z = (botSurPts(0).z + topSurPts(1).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point1_", ObjPoint
                        Set ObjPoint = Nothing
                        
                        'Right FacePoint
                        x = (botSurPts(3).x + topSurPts(2).x) / 2
                        y = (botSurPts(3).y + topSurPts(2).y) / 2
                        z = (botSurPts(3).z + topSurPts(2).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point2_", ObjPoint
                        Set ObjPoint = Nothing
                         
                        'Front FacePoint
                        x = (botSurPts(0).x + topSurPts(3).x) / 2
                        y = (botSurPts(0).y + topSurPts(3).y) / 2
                        z = (botSurPts(0).z + topSurPts(3).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point3_", ObjPoint
                        Set ObjPoint = Nothing
                        
                        'Back FacePoint
                        x = (botSurPts(1).x + topSurPts(2).x) / 2
                        y = (botSurPts(1).y + topSurPts(2).y) / 2
                        z = (botSurPts(1).z + topSurPts(2).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point4_", ObjPoint
                        Set ObjPoint = Nothing
                        
                        'Bottom FacePoint
                        x = (botSurPts(0).x + botSurPts(2).x) / 2
                        y = (botSurPts(0).y + botSurPts(2).y) / 2
                        z = (botSurPts(0).z + botSurPts(2).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point5_", ObjPoint
                        Set ObjPoint = Nothing
                        
                        'Top FacePoint
                        x = (topSurPts(0).x + topSurPts(2).x) / 2
                        y = (topSurPts(0).y + topSurPts(2).y) / 2
                        z = (topSurPts(0).z + topSurPts(2).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point6_", ObjPoint
                        Set ObjPoint = Nothing

            Next iCount
            
            For iCount = 0 To 3
                Set topSurPts(iCount) = Nothing
                Set botSurPts(iCount) = Nothing
            Next iCount
            Set oGeomFactory = Nothing
    End If
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub



