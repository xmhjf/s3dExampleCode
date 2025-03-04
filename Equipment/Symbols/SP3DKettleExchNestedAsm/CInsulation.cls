VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:13:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Mar 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Kettle Exchanger Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E307 Symbol in Page no 304.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

'   Define front end types
    Const FrontEndType_A = 1
    Const FrontEndType_B = 2
    Const FrontEndType_C = 3
    Const FrontEndType_N = 5
    Const FrontEndType_D = 4
    Const FrontEndType_Q = 6

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjExBodyInsul As Object
    Dim ObjExTapBodyInsul As Object
    Dim ObjExneckBodyInsul As Object

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
    Dim parFrontEndType As Long
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)           'P1
    parExchangerNeckLength = arrayOfInputs(3)       'P2
    parExchangerNeckDiameter = arrayOfInputs(4)     'P3
    parExchangerTaperLength = arrayOfInputs(5)      'P4
    parExchangerDiameter = arrayOfInputs(6)         'P5
    parBundleFlangeTk = arrayOfInputs(7)            'P6
    parBundleFlangeDia = arrayOfInputs(8)           'P7
    parExchangerFlangeTk1 = arrayOfInputs(9)        'P8
    parBundlePullingLength = arrayOfInputs(10)      'P9
    parBotSupportCenFromPP = arrayOfInputs(11)      'P10
    parBottomSupportCentoCen = arrayOfInputs(12)    'P11
    parSupport1Thickness = arrayOfInputs(13)        'P12
    parSupport2Thickness = arrayOfInputs(14)        'P13
    parBottomSupportHeight = arrayOfInputs(15)      'P14
    parSupportLength = arrayOfInputs(16)            'P15
    parFrontEndFlangeDia = arrayOfInputs(17)        'P30
    parFrontEndFlangeTk1 = arrayOfInputs(18)        'P31
    parFrontEndLength1 = arrayOfInputs(19)          'P32
    parFrontEndLength2 = arrayOfInputs(20)          'P33 of type A/C/D/N or B
    parFrontEndFlangeTk2 = arrayOfInputs(21)        'P34
    parFrontEndFlangeTk3 = arrayOfInputs(22)        'P35
    parFrontEndDiameter = arrayOfInputs(23)       'P33 of type Q End
    parFrontEndType = arrayOfInputs(24)
    parInsulationThickness = arrayOfInputs(25)

    iOutput = 0

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

 ' Insert your code for output 10(ExchangerBody Insulation)
    'Place Elliptical Arc on Right Side
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double

    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpArc As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d
    Dim EleCollection           As Collection

    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    Dim Linepts(0 To 5) As Double

    cenX = parFrontEndLength1 + parExchangerLength
    cenY = 0
    cenZ = parExchangerDiameter / 2 - parExchangerNeckDiameter / 2

    MajorX = 0
    MajorY = 0
    MajorZ = parExchangerDiameter / 2 + parInsulationThickness

    mMRatio = (parExchangerDiameter / 4 + parInsulationThickness) / _
                    (parExchangerDiameter / 2 + parInsulationThickness)
    StartAngle = 1.5 * PI
    SweepAngle = PI / 2

    norX = 0
    norY = -1
    norZ = 0

    Set oEqpArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    'Horizontal Line
    Linepts(0) = parFrontEndLength1 + parExchangerLength
    Linepts(1) = 0
    Linepts(2) = parExchangerDiameter / 2 + parInsulationThickness + _
                        (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    Linepts(3) = parFrontEndLength1 + parExchangerNeckLength + parExchangerTaperLength
    Linepts(4) = 0
    Linepts(5) = Linepts(2)

    Set oEqpLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Set EleCollection = New Collection
    EleCollection.Add oEqpArc
    EleCollection.Add oEqpLine

    startPoint.Set parFrontEndLength1 + parExchangerLength + parExchangerDiameter / 4 + parInsulationThickness, _
                            0, _
                            (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVect.Set 1, 0, 0
    revCenPt.Set (Linepts(0) + Linepts(3)) / 2, 0, _
                            (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    'Revolve it about X-Axiz
    Set ObjExBodyInsul = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExBodyInsul
    Set ObjExBodyInsul = Nothing

 ' Insert your code for output 11(ExchangerTaperBody ins)
    Dim topCenter As New AutoMath.DPosition
    Dim baseCenter As New AutoMath.DPosition
    Dim obaseCircle As IngrGeom3D.Circle3d
    Dim otopCircle As IngrGeom3D.Circle3d
    'P2+P32
    baseCenter.Set (parExchangerNeckLength + parFrontEndLength1), _
                                0, _
                                0
    'P32+P2+P4
    topCenter.Set (parExchangerNeckLength + parExchangerTaperLength + parFrontEndLength1), _
                            0, _
                            parExchangerDiameter / 2 - parExchangerNeckDiameter / 2

    Set obaseCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, baseCenter.x, _
                                    baseCenter.y, baseCenter.z, 1, 0, 0, parExchangerNeckDiameter / 2 + parInsulationThickness)

    Set otopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, topCenter.x, _
                                    topCenter.y, topCenter.z, 1, 0, 0, parExchangerDiameter / 2 + parInsulationThickness)

    Set ObjExTapBodyInsul = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            obaseCircle, otopCircle, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExTapBodyInsul
    Set ObjExTapBodyInsul = Nothing
    Set obaseCircle = Nothing
    Set otopCircle = Nothing

 ' Insert your code for output 12(Exchanger Neck Portion insulation)
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 20) As Double

    'Point 1
    dblStrPts(0) = (parExchangerNeckLength + parFrontEndLength1)
    dblStrPts(1) = 0
    dblStrPts(2) = parExchangerNeckDiameter / 2 + parInsulationThickness
    'Point2             P32+P6+P8 + parInsulationThickness
    dblStrPts(3) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1 + parInsulationThickness)
    dblStrPts(4) = 0
    dblStrPts(5) = dblStrPts(2)
    'Point3
    dblStrPts(6) = dblStrPts(3)
    dblStrPts(7) = 0
    dblStrPts(8) = parFrontEndFlangeDia / 2 + parInsulationThickness
    'Point4
    dblStrPts(9) = (parFrontEndLength1 + parBundleFlangeTk)
    dblStrPts(10) = 0
    dblStrPts(11) = dblStrPts(8)
    'Point5
    dblStrPts(12) = dblStrPts(9)
    dblStrPts(13) = 0
    dblStrPts(14) = parBundleFlangeDia / 2 + parInsulationThickness
    'Point6
    dblStrPts(15) = (parFrontEndLength1)
    dblStrPts(16) = 0
    dblStrPts(17) = dblStrPts(14)
    'Point7
    dblStrPts(18) = (parFrontEndLength1)
    dblStrPts(19) = 0
    dblStrPts(20) = 0

    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dblStrPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (dblStrPts(18) + dblStrPts(0)) / 2, _
                                0, _
                                0
    Set ObjExneckBodyInsul = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExneckBodyInsul
    Set ObjExneckBodyInsul = Nothing
    Set oLineStr = Nothing

    
    Set oGeomFactory = Nothing
    Set startPoint = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing
    Set EleCollection = Nothing
    Set topCenter = Nothing
    Set baseCenter = Nothing
    Set revVector = Nothing
    Set revCenterPt = Nothing
    Set oEqpArc = Nothing
    Set oEqpLine = Nothing
    Set oEqpComplexStr = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
