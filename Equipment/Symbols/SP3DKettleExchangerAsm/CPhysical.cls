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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:06 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Apr 18 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is E307 Equipment Assembly Kettle Exchanger symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'
'   4 Sept 2003     Nabil Nicolas           Define Support bottom planes with downwards vectors;
'                                           Define Support1 as box to provide planes for constraints;
'                                           Define only one Default Surface at Support2; dissociate support height from square root calculations based on support length.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjExchangerBody As Object
    Dim ObjExTaperBody As Object
    Dim ObjExneckBody As Object
    Dim ObjExFrontEnd As Object
    Dim ObjExchangerSup1 As Object
    Dim ObjExchangerSup2 As Object
    Dim ObjDefSurface1 As Object
    Dim ObjDefSurface2 As Object
        
    Dim parExchangerLength  As Double
    Dim parExchangerNeckLength As Double
    Dim parExchangerNeckDiameter As Double
    Dim parExchangerTaperLength As Double
    Dim parExchangerDiameter  As Double
    Dim parBundleFlangeTk  As Double
    Dim parBundleFlangeDia  As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parBundlePullingLength  As Double
    Dim parBotSupportCenFromPP  As Double
    Dim parBottomSupportCentoCen  As Double
    Dim parSupport1Thickness  As Double
    Dim parSupport2Thickness  As Double
    Dim parBottomSupportHeight  As Double
    Dim parSupportLength As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1  As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)                           'P1
    parExchangerNeckLength = arrayOfInputs(3)                   'P2
    parExchangerNeckDiameter = arrayOfInputs(4)                 'P3
    parExchangerTaperLength = arrayOfInputs(5)                  'P4
    parExchangerDiameter = arrayOfInputs(6)                         'P5
    parBundleFlangeTk = arrayOfInputs(7)                                'P6
    parBundleFlangeDia = arrayOfInputs(8)                               'P7
    parExchangerFlangeTk1 = arrayOfInputs(9)                    'P8
    parBundlePullingLength = arrayOfInputs(10)                  'P9
    parBotSupportCenFromPP = arrayOfInputs(11)              'P10
    parBottomSupportCentoCen = arrayOfInputs(12)        'P11
    parSupport1Thickness = arrayOfInputs(13)                    'P12
    parSupport2Thickness = arrayOfInputs(14)                    'P13
    parBottomSupportHeight = arrayOfInputs(15)              'P14
    parSupportLength = arrayOfInputs(16)                        'P15
    parFrontEndFlangeDia = arrayOfInputs(17)                'P30
    parFrontEndFlangeTk1 = arrayOfInputs(18)                'P31
    parFrontEndLength1 = arrayOfInputs(19)                  'P32
    parFrontEndLength2 = arrayOfInputs(20)                  'P33
    parFrontEndFlangeTk2 = arrayOfInputs(21)                'P34
    parFrontEndFlangeTk3 = arrayOfInputs(22)                'P35
    parInsulationThickness = arrayOfInputs(23)
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    iOutput = 0
 ' Insert your code for output 1(ExchangerBody)
    'Place Elliptical Arc on Right Side
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim majorX As Double, majorY As Double, majorZ As Double
    Dim mMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    
    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpArc As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d
    Dim EleCollection           As Collection
    
    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    Dim Linepts(0 To 5) As Double
    
    cenX = parFrontEndLength1 + parExchangerLength
    cenY = CenterPos.y
    cenZ = parExchangerDiameter / 2 - parExchangerNeckDiameter / 2
    
    majorX = CenterPos.x
    majorY = CenterPos.y
    majorZ = parExchangerDiameter / 2
    
    mMRatio = 0.5
    StartAngle = 1.5 * PI
    SweepAngle = PI / 2
    
    norX = 0
    norY = -1
    norZ = 0
    
    Set oEqpArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    'Horizontal Line
    Linepts(0) = parFrontEndLength1 + parExchangerLength
    Linepts(1) = CenterPos.y
    Linepts(2) = parExchangerDiameter / 2 + _
                        (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    Linepts(3) = parFrontEndLength1 + parExchangerNeckLength + parExchangerTaperLength
    Linepts(4) = CenterPos.y
    Linepts(5) = Linepts(2)
    
    Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Set EleCollection = New Collection
    EleCollection.Add oEqpArc
    EleCollection.Add oEqpLine
    
    startPoint.Set parFrontEndLength1 + parExchangerLength + parExchangerDiameter / 4, _
                            CenterPos.y, _
                            (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVect.Set 1, 0, 0
    revCenPt.Set (Linepts(0) + Linepts(3)) / 2, CenterPos.y, _
                            (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    'Revolve it about X-Axiz
    Set ObjExchangerBody = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, PI * 2, True)
' Set the output
  
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
 
    Set axisVect = Nothing
    Set revCenPt = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oEqpComplexStr
    Objcurves.Remove
    
    Set Objcurves = Nothing
    Set oEqpLine = Nothing
    Set oEqpArc = Nothing
    
    Dim count As Integer
    For count = 1 To EleCollection.count
        EleCollection.Remove 1
    Next count
    Set EleCollection = Nothing
    Set oEqpComplexStr = Nothing
 
 ' Insert your code for output 2(ExchangerTaperBody)
    Dim topCenter As New AutoMath.DPosition
    Dim baseCenter As New AutoMath.DPosition
    Dim obaseCircle As IngrGeom3D.Circle3d
    Dim otopCircle As IngrGeom3D.Circle3d
    'P2+P32
    baseCenter.Set (parExchangerNeckLength + parFrontEndLength1), _
                                CenterPos.y, _
                                CenterPos.z
    'P32+P2+P4
    topCenter.Set (parExchangerNeckLength + parExchangerTaperLength + parFrontEndLength1), _
                            CenterPos.y, _
                            parExchangerDiameter / 2 - parExchangerNeckDiameter / 2
    Set obaseCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, baseCenter.x, _
                                            baseCenter.y, baseCenter.z, 1, 0, 0, parExchangerNeckDiameter / 2)
    
    Set otopCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, topCenter.x, _
                                            topCenter.y, topCenter.z, 1, 0, 0, parExchangerDiameter / 2)

    Set ObjExTaperBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            obaseCircle, otopCircle, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExTaperBody
    Set ObjExTaperBody = Nothing
    Set obaseCircle = Nothing
    Set otopCircle = Nothing
    
 ' Insert your code for output 3(Exchanger Neck Portion)
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 20) As Double

    'Point 1
    dblStrPts(0) = (parExchangerNeckLength + parFrontEndLength1)
    dblStrPts(1) = CenterPos.y
    dblStrPts(2) = parExchangerNeckDiameter / 2
    'Point2             P32+P6+P8
    dblStrPts(3) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    dblStrPts(4) = CenterPos.y
    dblStrPts(5) = dblStrPts(2)
    'Point3
    dblStrPts(6) = dblStrPts(3)
    dblStrPts(7) = CenterPos.y
    dblStrPts(8) = parFrontEndFlangeDia / 2
    'Point4
    dblStrPts(9) = (parFrontEndLength1 + parBundleFlangeTk)
    dblStrPts(10) = CenterPos.y
    dblStrPts(11) = dblStrPts(8)
    'Point5
    dblStrPts(12) = dblStrPts(9)
    dblStrPts(13) = CenterPos.y
    dblStrPts(14) = parBundleFlangeDia / 2
    'Point6
    dblStrPts(15) = (parFrontEndLength1)
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = dblStrPts(14)
    'Point7
    dblStrPts(18) = (parFrontEndLength1)
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = CenterPos.z
    
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dblStrPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (dblStrPts(18) + dblStrPts(0)) / 2, _
                                CenterPos.y, _
                                CenterPos.z

    Set ObjExneckBody = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExneckBody
    Set ObjExneckBody = Nothing
    Set oLineStr = Nothing
    
 ' Insert your code for output 4(Exchanger Front End)
    'Create a Point String and Revolve it
    Dim oEndLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 26) As Double
    
    'Point 1
    dblEndPts(0) = parFrontEndLength1           '+P32
    dblEndPts(1) = CenterPos.y
    dblEndPts(2) = CenterPos.z
    'Point 2
    dblEndPts(3) = parFrontEndLength1
    dblEndPts(4) = CenterPos.y
    dblEndPts(5) = parFrontEndFlangeDia / 2     'P30/2
    'Point 3
    dblEndPts(6) = (parFrontEndLength1 - parFrontEndFlangeTk1)
    dblEndPts(7) = CenterPos.y
    dblEndPts(8) = dblEndPts(5)                 'P30/2
    'Point 4
    dblEndPts(9) = dblEndPts(6)
    dblEndPts(10) = CenterPos.y
    dblEndPts(11) = parExchangerNeckDiameter / 2              '(P30-P2)/2
    'Point 5
    dblEndPts(12) = -(parFrontEndLength2 - parFrontEndFlangeTk3)
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = dblEndPts(11)
    'Point 6
    dblEndPts(15) = dblEndPts(12)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parFrontEndFlangeDia / 2
    'Point 7
    dblEndPts(18) = -parFrontEndLength2
    dblEndPts(19) = CenterPos.y
    dblEndPts(20) = dblEndPts(17)
    'Point 8
    dblEndPts(21) = -(parFrontEndLength2 + parFrontEndFlangeTk2)
    dblEndPts(22) = CenterPos.y
    dblEndPts(23) = dblEndPts(17)
    'Point 9
    dblEndPts(24) = dblEndPts(21)
    dblEndPts(25) = CenterPos.y
    dblEndPts(26) = CenterPos.z

    Set oEndLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set 0, 0, 0
    Set ObjExFrontEnd = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExFrontEnd
    Set ObjExFrontEnd = Nothing
    Set oEndLineStr = Nothing
    
 ' Insert your code for output 5(Exchanger Support)
    Dim oSupLine As IngrGeom3D.LineString3d
    Dim dblSupPts(0 To 14) As Double
    Dim projVect As New AutoMath.DVector
    Dim projHt As Double, ActualProj As Double
    Dim Sup1L As IJDPosition
    Dim Sup1H As IJDPosition
    Set Sup1L = New DPosition
    Set Sup1H = New DPosition
    Dim SupLZ As Double
    If (parBottomSupportHeight > parExchangerNeckDiameter / 2 And parSupportLength > 0) Then
        If parSupport1Thickness > 0 Then
'            'Point 1        x (P10-P13/2)
'            dblSupPts(0) = (parBotSupportCenFromPP - parSupport1Thickness / 2)
'            dblSupPts(1) = -(parSupportLength / 2)
'            dblSupPts(2) = -parBottomSupportHeight
'            'Point2         x (P10+P13/2)*cos
'            dblSupPts(3) = dblSupPts(0) + parSupport1Thickness
'            dblSupPts(4) = dblSupPts(1)
'            dblSupPts(5) = -parBottomSupportHeight
'            'Point3
'            dblSupPts(6) = dblSupPts(3)
'            dblSupPts(7) = (parSupportLength / 2)
'            dblSupPts(8) = -parBottomSupportHeight
'            'Point4
'            dblSupPts(9) = dblSupPts(0)
'            dblSupPts(10) = dblSupPts(7)
'            dblSupPts(11) = -parBottomSupportHeight
'            'Point5
'            dblSupPts(12) = dblSupPts(0)
'            dblSupPts(13) = dblSupPts(1)
'            dblSupPts(14) = dblSupPts(2)
            SupLZ = cenZ - parBottomSupportHeight
            'Point 1        x (P10-P13/2)
            dblSupPts(0) = (parBotSupportCenFromPP - parSupport1Thickness / 2)
            dblSupPts(1) = -(parSupportLength / 2)
            dblSupPts(2) = SupLZ
            'Point2
            dblSupPts(3) = dblSupPts(0)
            dblSupPts(4) = -dblSupPts(1)
            dblSupPts(5) = SupLZ
            'Point3
            dblSupPts(6) = (parBotSupportCenFromPP + parSupport1Thickness / 2)
            dblSupPts(7) = (parSupportLength / 2)
            dblSupPts(8) = SupLZ
            'Point4         x (P10+P13/2)*cos
            dblSupPts(9) = dblSupPts(0) + parSupport1Thickness
            dblSupPts(10) = dblSupPts(1)
            dblSupPts(11) = SupLZ
            'Point5
            dblSupPts(12) = dblSupPts(0)
            dblSupPts(13) = dblSupPts(1)
            dblSupPts(14) = dblSupPts(2)
            
'            projHt = Sqr((parExchangerDiameter / 2) * (parExchangerDiameter / 2) - (parSupportLength / 2) * (parSupportLength / 2))
'            projVect.Set 0, 0, 1
'            ActualProj = parBottomSupportHeight - projHt
'            Set oSupLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)

            Set ObjDefSurface1 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblSupPts)
'            Set ObjExchangerSup1 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, False)
            
            Sup1L.Set dblSupPts(0), dblSupPts(1), dblSupPts(2)
            Sup1H.Set dblSupPts(6), dblSupPts(7), 0
            Set ObjExchangerSup1 = PlaceBox(m_OutputColl, Sup1L, Sup1H)
            
            ' Set the output
            m_OutputColl.AddOutput "Support1", ObjExchangerSup1
            Set ObjExchangerSup1 = Nothing
            
            ' Set the output' Insert your code for output 7(Default Surface 1)
            m_OutputColl.AddOutput "Support1BotPlane", ObjDefSurface1
            Set ObjDefSurface1 = Nothing
            Set oSupLine = Nothing
        End If
        ' Insert your code for output 6(Exchanger Support)
        
        If (parBottomSupportCentoCen > parSupport1Thickness And parSupport2Thickness > 0) Then

'            'Point1
'            dblSupPts(0) = (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2)
'            dblSupPts(1) = -(parSupportLength / 2)
'            dblSupPts(2) = -parBottomSupportHeight
'            'Point2
'            dblSupPts(3) = dblSupPts(0) + parSupport2Thickness
'            dblSupPts(4) = dblSupPts(1)
'            dblSupPts(5) = -parBottomSupportHeight
'            'Point3
'            dblSupPts(6) = dblSupPts(3)
'            dblSupPts(7) = (parSupportLength / 2)
'            dblSupPts(8) = -parBottomSupportHeight
'            'Point4
'            dblSupPts(9) = dblSupPts(0)
'            dblSupPts(10) = dblSupPts(7)
'            dblSupPts(11) = -parBottomSupportHeight
'            'Point5
'            dblSupPts(12) = dblSupPts(0)
'            dblSupPts(13) = dblSupPts(1)
'            dblSupPts(14) = dblSupPts(2)

            'Point1
            dblSupPts(0) = (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2)
            dblSupPts(1) = -(parSupportLength / 2)
            dblSupPts(2) = SupLZ
            'Point2
            dblSupPts(3) = dblSupPts(0)
            dblSupPts(4) = -dblSupPts(1)
            dblSupPts(5) = SupLZ
             'Point3
            dblSupPts(6) = (parBotSupportCenFromPP + parBottomSupportCentoCen + parSupport2Thickness / 2)
            dblSupPts(7) = (parSupportLength / 2)
            dblSupPts(8) = SupLZ
            'Point4
            dblSupPts(9) = dblSupPts(0) + parSupport2Thickness
            dblSupPts(10) = dblSupPts(1)
            dblSupPts(11) = SupLZ
            'Point5
            dblSupPts(12) = dblSupPts(0)
            dblSupPts(13) = dblSupPts(1)
            dblSupPts(14) = dblSupPts(2)

'            projHt = Sqr((parExchangerDiameter / 2) * (parExchangerDiameter / 2) - (parSupportLength / 2) * (parSupportLength / 2))
            projVect.Set 0, 0, 1
            ActualProj = -SupLZ

            Set oSupLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
            Set ObjDefSurface2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblSupPts)
            Set ObjExchangerSup2 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, False)
             
            ' Set the output
            m_OutputColl.AddOutput "Support2", ObjExchangerSup2
            Set ObjExchangerSup2 = Nothing
            
            ' Set the output
            m_OutputColl.AddOutput "DefaultSurface", ObjDefSurface2
            Set ObjDefSurface2 = Nothing
            Set oSupLine = Nothing
        End If
    End If
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.05)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "KettleExchangerControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
 
End Sub
