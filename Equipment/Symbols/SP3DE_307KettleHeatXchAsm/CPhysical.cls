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
'   Creation Date:  Monday,March 5 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Twelve Outputs:
'      In which Default Surface is created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     
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
    
    'Define Front End types
    Const FrontEndType_A = 1
    Const FrontEndType_B = 2
    Const FrontEndType_C = 3
    Const FrontEndType_D = 4
    Const FrontEndType_N = 5
    Const FrontEndType_Q = 6

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
    Dim ObjExchangerBody As Object
    Dim ObjExTaperBody As Object
    Dim ObjExneckBody As Object
    Dim ObjExchangerFlange As Object
    Dim ObjBundleFlange As Object
    Dim ObjExchangerSup1 As Object
    Dim ObjExchangerSup2 As Object
    
    Dim parFrontEndType As Long
    Dim parVesselHeadType As Long
    Dim parExchangerLength As Double
    Dim parExchangerNeckLength As Double
    Dim parExchangerNeckDiameter As Double
    Dim parExchangerTaperLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parBundlePullingLength As Double
    Dim parBotSupportCenFromPP As Double
    Dim parBottomSupportCentoCen As Double
    Dim parSupport1Thickness As Double
    Dim parSupport2Thickness As Double
    Dim parBottomSupportHeight As Double
    Dim parSupportLength As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parFrontEndDiameter As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    Dim parFP1XBoltHole As Double
    Dim parFP1YBoltHole As Double
    Dim parFP2XBoltHole As Double
    Dim parFP2YBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndType = arrayOfInputs(2)
    parVesselHeadType = arrayOfInputs(3)
    parExchangerLength = arrayOfInputs(4)           'P1
    parExchangerNeckLength = arrayOfInputs(5)       'P2
    parExchangerNeckDiameter = arrayOfInputs(6)     'P3
    parExchangerTaperLength = arrayOfInputs(7)      'P4
    parExchangerDiameter = arrayOfInputs(8)         'P5
    parExchangerFlangeTk1 = arrayOfInputs(9)        'P8
    parBundleFlangeDia = arrayOfInputs(10)          'P7
    parBundleFlangeTk = arrayOfInputs(11)           'P6
    parBundlePullingLength = arrayOfInputs(12)      'P9
    parSupportLength = arrayOfInputs(13)            'P15
    parBotSupportCenFromPP = arrayOfInputs(14)      'P10
    parBottomSupportCentoCen = arrayOfInputs(15)    'P11
    parBottomSupportHeight = arrayOfInputs(16)      'P14
    parSupport1Thickness = arrayOfInputs(17)        'P12
    parSupport2Thickness = arrayOfInputs(18)        'P13
    parFrontEndFlangeDia = arrayOfInputs(19)        'P30
    parFrontEndFlangeTk1 = arrayOfInputs(20)        'P31
    parFrontEndFlangeTk2 = arrayOfInputs(21)        'P34
    parFrontEndFlangeTk3 = arrayOfInputs(22)        'P35
    parFrontEndLength1 = arrayOfInputs(23)          'P32
    parFrontEndLength2 = arrayOfInputs(24)          'P33
    parFrontEndDiameter = arrayOfInputs(25)         'P33 of type Q End
    parEndHeadConeTopDiameter = arrayOfInputs(26)   'P51
    parEndHeadConeHeight = arrayOfInputs(27)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(28)     'P51 for TORC & TORS
    parEndHeadDomeradius = arrayOfInputs(29)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(30)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(31)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(32)   'P50 for TORS
    parInsulationThickness = arrayOfInputs(33)
    parFP1XBoltHole = arrayOfInputs(34)
    parFP1YBoltHole = arrayOfInputs(35)
    parFP2XBoltHole = arrayOfInputs(36)
    parFP2YBoltHole = arrayOfInputs(37)
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    'Insert your code for output (DefaultSurface)
    'Surface will be at Origin
    'Create non-persistent circle to use for creating default surface ---
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                        0, 0, 0, _
                                                        0, 0, -1, _
                                                        0.1 * parExchangerDiameter / 2)
    
    'Create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objCircle = Nothing
    
    'Origin is taken on axis of Exchanger End, at the junction of P32 and P33 dimensions.
    'Insert your code for ExchangerBody
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    stPoint.Set parFrontEndLength1 + parExchangerLength, 0, _
                    (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    enPoint.Set parFrontEndLength1 + parExchangerNeckLength + parExchangerTaperLength, 0, _
                    (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)

    Set ObjExchangerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parExchangerDiameter, False)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
 
    'Insert your code for ExchangerTaperBody
    Dim topCenter As New AutoMath.DPosition
    Dim baseCenter As New AutoMath.DPosition
    Dim oBaseCircle As IngrGeom3D.Circle3d
    Dim oTopCircle As IngrGeom3D.Circle3d

    'X-coordinate: P2+P32
    baseCenter.Set (parExchangerNeckLength + parFrontEndLength1), 0, 0
    'X-coordinate: P32+P2+P4
    topCenter.Set (parExchangerNeckLength + parExchangerTaperLength + parFrontEndLength1), _
                            0, _
                            parExchangerDiameter / 2 - parExchangerNeckDiameter / 2
    
    Set oBaseCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, baseCenter.x, _
                                            baseCenter.y, baseCenter.z, 1, 0, 0, parExchangerNeckDiameter / 2)
    Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, topCenter.x, _
                                            topCenter.y, topCenter.z, 1, 0, 0, parExchangerDiameter / 2)
    Set ObjExTaperBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oBaseCircle, oTopCircle, False)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExTaperBody
    Set ObjExTaperBody = Nothing
    Set oBaseCircle = Nothing
    Set oTopCircle = Nothing
    
    'Insert your code for Exchanger Neck Portion
    'X-coordinate: P32 + P2
    stPoint.Set (parExchangerNeckLength + parFrontEndLength1), 0, 0
    'X-coordinate: P32+P6+P8
    enPoint.Set (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1), _
                0, 0

    Set ObjExneckBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parExchangerNeckDiameter, False)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExneckBody
    Set ObjExneckBody = Nothing

    'Insert your code for ExchangerFlange
    'X-coordinate: P32+P6+P8
    stPoint.Set (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1), 0, 0
    'X-coordinate: P32+P6
    enPoint.Set (parFrontEndLength1 + parBundleFlangeTk), 0, 0
                
    Set ObjExchangerFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parFrontEndFlangeDia, False)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerFlange
    Set ObjExchangerFlange = Nothing

   'Insert your code for BundleFlange
   'X-coordinate: P32+P6
    stPoint.Set (parFrontEndLength1 + parBundleFlangeTk), 0, 0
   'X-coordinate: P32
    enPoint.Set parFrontEndLength1, 0, 0
                
    Set ObjBundleFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parBundleFlangeDia, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundleFlange
    Set ObjBundleFlange = Nothing
    
    'Insert your code for Exchanger Support
    Dim oSupLine As IngrGeom3D.LineString3d
    Dim dblSupPts(0 To 14) As Double
    Dim projVect As New AutoMath.DVector
    Dim ActualProj As Double
    Dim Sup1L As IJDPosition
    Dim Sup1H As IJDPosition
    Set Sup1L = New DPosition
    Set Sup1H = New DPosition
    Dim SupLZ As Double
    
    If CmpDblGreaterthan(parBottomSupportHeight, parExchangerNeckDiameter / 2) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
    'If (parBottomSupportHeight > parExchangerNeckDiameter / 2 And parSupportLength > 0) Then
        'If parSupport1Thickness > 0 Then
        If CmpDblGreaterthan(parSupport1Thickness, LINEAR_TOLERANCE) Then
            SupLZ = parBottomSupportHeight - (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
            Sup1L.Set (parBotSupportCenFromPP - parSupport1Thickness / 2), _
                        -(parSupportLength / 2), -SupLZ

            Sup1H.Set (parBotSupportCenFromPP + parSupport1Thickness / 2), _
                      (parSupportLength / 2), 0
            Set ObjExchangerSup1 = PlaceBox(m_OutputColl, Sup1L, Sup1H)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Support1", ObjExchangerSup1
            Set ObjExchangerSup1 = Nothing
        End If
        'Insert your code for Exchanger Support2
        'If (parBottomSupportCentoCen > parSupport1Thickness And parSupport2Thickness > 0) Then
        If CmpDblGreaterthan(parBottomSupportCentoCen, parSupport1Thickness) And CmpDblGreaterthan(parSupport2Thickness, LINEAR_TOLERANCE) Then
            'Point1
            dblSupPts(0) = (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2)
            dblSupPts(1) = -(parSupportLength / 2)
            dblSupPts(2) = -SupLZ
            'Point2
            dblSupPts(3) = dblSupPts(0)
            dblSupPts(4) = parSupportLength / 2
            dblSupPts(5) = -SupLZ
             'Point3
            dblSupPts(6) = (parBotSupportCenFromPP + parBottomSupportCentoCen + parSupport2Thickness / 2)
            dblSupPts(7) = (parSupportLength / 2)
            dblSupPts(8) = -SupLZ
            'Point4
            dblSupPts(9) = dblSupPts(6)
            dblSupPts(10) = dblSupPts(1)
            dblSupPts(11) = -SupLZ
            'Point5
            dblSupPts(12) = dblSupPts(0)
            dblSupPts(13) = dblSupPts(1)
            dblSupPts(14) = dblSupPts(2)

            projVect.Set 0, 0, 1
            ActualProj = SupLZ
            Set oSupLine = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
            Set ObjExchangerSup2 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Support2", ObjExchangerSup2
            Set ObjExchangerSup2 = Nothing
        End If
    End If
    
    'Inser your code for creating the edges for Support(s)
    Dim iCount As Integer
    Dim jCount As Integer
    Dim Points(0 To 1, 0 To 23) As Double
    Dim ObjEdgeColl As New Collection
    Dim Center As New AutoMath.DPosition
    Dim SupportThickness As Double
    If CmpDblGreaterthan(parBottomSupportHeight, parExchangerNeckDiameter / 2) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        If CmpDblGreaterthan(parSupport1Thickness, LINEAR_TOLERANCE) Or _
            CmpDblGreaterthan(parBottomSupportCentoCen, parSupport1Thickness) And CmpDblGreaterthan(parSupport2Thickness, LINEAR_TOLERANCE) Then
    
            SupportThickness = parSupport1Thickness
            Center.Set parBotSupportCenFromPP - SupportThickness / 2, 0, -SupLZ / 2
            jCount = 0
            While jCount <= 1
                If CmpDblLessThanOrEqualTo(parSupport1Thickness, LINEAR_TOLERANCE) Then
                    Center.Set parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2, 0, -SupLZ / 2
                    SupportThickness = parSupport2Thickness
                    jCount = 1
                End If
                For iCount = 0 To 12 Step 12
                    Points(jCount, 0 + iCount) = Center.x
                    Points(jCount, 1 + iCount) = Center.y + parSupportLength / 2
                    Points(jCount, 2 + iCount) = Center.z + SupLZ / 2
            
                    Points(jCount, 3 + iCount) = Center.x
                    Points(jCount, 4 + iCount) = Center.y + parSupportLength / 2
                    Points(jCount, 5 + iCount) = Center.z - SupLZ / 2
                    
                    Points(jCount, 6 + iCount) = Center.x
                    Points(jCount, 7 + iCount) = Center.y - parSupportLength / 2
                    Points(jCount, 8 + iCount) = Center.z + SupLZ / 2
                    
                    Points(jCount, 9 + iCount) = Center.x
                    Points(jCount, 10 + iCount) = Center.y - parSupportLength / 2
                    Points(jCount, 11 + iCount) = Center.z - SupLZ / 2
                    
                    Center.x = Center.x + SupportThickness
                Next iCount
                If CmpDblGreaterthan(parBottomSupportCentoCen, parSupport1Thickness) And CmpDblGreaterthan(parSupport2Thickness, LINEAR_TOLERANCE) _
                        And jCount = 0 Then
                    Center.Set parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2, 0, -SupLZ / 2
                    SupportThickness = parSupport2Thickness
                 Else
                     jCount = jCount + 1
                End If
                    jCount = jCount + 1
            Wend
            'Edges
            'Edge1 Point1 to Point2
            jCount = 0
            While jCount <= 1
                If CmpDblLessThanOrEqualTo(parSupport1Thickness, LINEAR_TOLERANCE) Then
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
                If Not CmpDblGreaterthan(parBottomSupportCentoCen, parSupport1Thickness) And CmpDblGreaterthan(parSupport2Thickness, LINEAR_TOLERANCE) Then
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
            Set Center = Nothing
            Set oGeomFactory = Nothing
            Set stPoint = Nothing
            Set enPoint = Nothing
            Set topCenter = Nothing
            Set baseCenter = Nothing
            Set projVect = Nothing
            Set Sup1L = Nothing
            Set Sup1H = Nothing
            Set oSupLine = Nothing
            Set Center = Nothing
        End If
    End If
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Coordinate system is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |       |
    '   |                  |       |
    '   |                  |       |
    '   -----> X           |       |
    '   Symbol CS          |       |
    '                      |       |
    '                      |   X   |
    '                      |   ^   |
    '                      |   |   |
    '                      |   |   |
    '                      |---|---|----> Y
    '                   Port CS

    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double

    'The origin of the port is taken to be at the corner point of the support base.
    dOrigin(0) = parBotSupportCenFromPP
    dOrigin(1) = 0
    dOrigin(2) = -(parBottomSupportHeight - (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2))
    
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

    holes(0, 1) = -parFP1XBoltHole
    holes(0, 2) = -parFP1YBoltHole
    holes(1, 1) = parFP1XBoltHole
    holes(1, 2) = -parFP1YBoltHole
    holes(2, 1) = parFP1XBoltHole
    holes(2, 2) = parFP1YBoltHole
    holes(3, 1) = -parFP1XBoltHole
    holes(3, 2) = parFP1YBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "EqpFoundationPort1", ObjFoundationPort
    Set ObjFoundationPort = Nothing

    'The origin of the port is taken to be at the corner point of the support base.
    dOrigin(0) = parBotSupportCenFromPP + parBottomSupportCentoCen
    dOrigin(1) = 0
    dOrigin(2) = -(parBottomSupportHeight - (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2))

    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort2", _
                                                    False, m_OutputColl.ResourceManager)
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parFP2XBoltHole
    holes(0, 2) = -parFP2YBoltHole
    holes(1, 1) = parFP2XBoltHole
    holes(1, 2) = -parFP2YBoltHole
    holes(2, 1) = parFP2XBoltHole
    holes(2, 2) = parFP2YBoltHole
    holes(3, 1) = -parFP2XBoltHole
    holes(3, 2) = parFP2YBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "EqpFoundationPort2", ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


