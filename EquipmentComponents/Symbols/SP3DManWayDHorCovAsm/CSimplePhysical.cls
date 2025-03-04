VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:54 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, March 212003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Vessel part, Davit for Horizontal Cover, is as per DOW chemical company document. This part is
'   constructed for 20" and 24" pipes. Refer DOW document reference G9G-7005-01 for more details.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Dim PI As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim parCentertoDavitCover As Double
    Dim parVesselDiameter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    iOutput = 0
   RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
' Place Nozzle 1
    Dim ObjManHole             As IJDNozzle
    Dim oDir As AutoMath.DVector
    Dim oPlacePoint As AutoMath.DPosition
    Dim nozzleExtension As Double
    Dim NozzleLength As Double
    
    
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition

    nozzleExtension = 0.2 * (parVesselDiameter / 2)
    NozzleLength = parCentertoDavitCover - parVesselDiameter / 2 + nozzleExtension
    oDir.Set 0, 0, 1
    oPlacePoint.Set 0, 0, 0
    
    Set ObjManHole = CreateNozzlePHWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjManHole
    Set ObjManHole = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
' Insert your code for output 2(Blind Flange)
    Dim flangeGap As Double
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    flangeGap = 0.004
    stPoint.Set 0, 0, flangeGap
    enPoint.Set 0, 0, flangeGap + flangeThick
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, flangeDiam
    
' Insert your code for output 3(Box over the blind flange)
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim dblPoints(0 To 14) As Double
    Dim cotterLength As Double
    Dim cotterWidth As Double
    Dim cotterDepth As Double
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim ProVect As AutoMath.DVector
    Dim CotterPinBox As Object
    
    Set ProVect = New AutoMath.DVector
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    cotterLength = 0.05         'As per DOW 0.5"+13/16"+0.5"
    cotterWidth = 0.04          'As per DOW 1 5/8", 40mm
    cotterDepth = 0.07          'As per DOW 1 5/8"+ 5/8"+5/8"
    
    'Point 1
    dblPoints(0) = -cotterLength / 2
    dblPoints(1) = -cotterWidth / 2
    dblPoints(2) = flangeGap + flangeThick
    'Point 2
    dblPoints(3) = cotterLength / 2
    dblPoints(4) = -cotterWidth / 2
    dblPoints(5) = flangeGap + flangeThick
    'Point 3
    dblPoints(6) = cotterLength / 2
    dblPoints(7) = cotterWidth / 2
    dblPoints(8) = flangeGap + flangeThick
    'Point4
    dblPoints(9) = -cotterLength / 2
    dblPoints(10) = cotterWidth / 2
    dblPoints(11) = flangeGap + flangeThick
    'Point5
    dblPoints(12) = -cotterLength / 2
    dblPoints(13) = -cotterWidth / 2
    dblPoints(14) = flangeGap + flangeThick
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set 0, 0, 1
    'Cotter Support
    Set CotterPinBox = PlaceProjection(m_OutputColl, oLineStr, ProVect, cotterDepth, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), CotterPinBox
    Set CotterPinBox = Nothing
    Set oLineStr = Nothing
    
' Insert your code for output 4(Vertical Rod above blind flange)
    Dim vertPinDia As Double
    Dim vertPinheight As Double
    Dim torusRadius As Double
    Dim davitPinDiameter As Double
    
    'As per DOW equipment davit pin diameter of 40mm for 150 pressure rating
    davitPinDiameter = 0.04
    
    
    torusRadius = 0.25
    vertPinDia = 0.02
    vertPinheight = torusRadius + davitPinDiameter
    stPoint.Set 0, 0, flangeGap + flangeThick
    enPoint.Set 0, 0, flangeGap + flangeThick + vertPinheight

'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, vertPinDia
    
' Insert your code for output 5(Horizontal portion of Connect rod)
    Dim topdavitSupWidth As Double
    Dim davitRotationAngle As Double
    Dim flangeEdgeTodavitCen As Double
    
    flangeEdgeTodavitCen = davitPinDiameter + 0.01
    davitRotationAngle = PI / 2
    'As per DOW document davit support width is sum of davit pin diameter, 10mm and 60mm
    topdavitSupWidth = davitPinDiameter + 0.01 + 0.06
    
    stPoint.Set -davitPinDiameter * Sin(davitRotationAngle), -davitPinDiameter * Cos(davitRotationAngle), _
                        flangeGap + flangeThick + vertPinheight - davitPinDiameter
    enPoint.Set Sin(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen - torusRadius), _
                        Cos(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen - torusRadius), _
                        stPoint.z

'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, davitPinDiameter

' Insert your code for output 6(Torus portion of Connecting rod )
    Dim centerX As Double
    Dim centerY As Double
    Dim centerZ As Double
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    Dim oCircle As IngrGeom3D.Circle3d
    Dim revVector As AutoMath.DVector
    Dim cenPoint As AutoMath.DPosition
    Dim ObjDavitTorusRod As Object
    
    Set revVector = New AutoMath.DVector
    Set cenPoint = New AutoMath.DPosition
    
        
    centerX = Sin(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen - torusRadius)
    centerY = Cos(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen - torusRadius)
    centerZ = flangeGap + flangeThick + vertPinheight - davitPinDiameter
    normalX = Sin(davitRotationAngle)
    normalY = Cos(davitRotationAngle)
    normalZ = 0
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, centerX, centerY, centerZ, _
                                                                    normalX, normalY, normalZ, davitPinDiameter / 2)
    revVector.Set Cos(davitRotationAngle), Sin(davitRotationAngle), 0
    cenPoint.Set Sin(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen - torusRadius), _
                            Cos(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen - torusRadius), _
                            centerZ - torusRadius
    Set ObjDavitTorusRod = PlaceRevolution(m_OutputColl, oCircle, revVector, cenPoint, PI / 2, True)

'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDavitTorusRod
    Set ObjDavitTorusRod = Nothing
    Set revVector = Nothing
    Set cenPoint = Nothing
    Set oCircle = Nothing
    
' Insert your code for output 7(Davit vertical rod in side)
    Dim topdavitSupTk1 As Double
    Dim davitSupsSpacing As Double
    Dim davitPinExtn As Double
    
    davitSupsSpacing = 0.15
    topdavitSupTk1 = 0.015
    davitPinExtn = 0.019
    stPoint.Set Sin(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen), _
                        Cos(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen), _
                        centerZ - torusRadius
    enPoint.Set stPoint.x, stPoint.y, _
                        -flangeThick / 3 - davitSupsSpacing - davitPinExtn
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, davitPinDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 8(Hexagonal Nut)
    Dim LineStrPoints(0 To 20) As Double
    Dim centerPoint As AutoMath.DPosition
    Dim hexNutDia As Double
    Dim angle As Double
    Dim nutThickness As Double
    Dim i As Integer
    Dim ObjHexNut As Object
    Set centerPoint = New AutoMath.DPosition
    
    hexNutDia = davitPinDiameter * 1.4
    nutThickness = 0.025
    
    centerPoint.Set Sin(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen), _
                              Cos(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen), _
                                -flangeThick / 3
    angle = PI / 6
    LineStrPoints(0) = centerPoint.x + (hexNutDia / 2) * Sin(angle)
    LineStrPoints(1) = centerPoint.y + (hexNutDia / 2) * Cos(angle)
    LineStrPoints(2) = centerPoint.z
    For i = 2 To 7
        angle = (PI / 6) + (PI / 3) * (i - 1)
        LineStrPoints(3 * i - 3) = centerPoint.x + (hexNutDia / 2) * Sin(angle)
        LineStrPoints(3 * i - 2) = centerPoint.y + (hexNutDia / 2) * Cos(angle)
        LineStrPoints(3 * i - 1) = centerPoint.z
    Next i
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    ProVect.Set 0, 0, 1
    
    'davit Support
    Set ObjHexNut = PlaceProjection(m_OutputColl, oLineStr, ProVect, nutThickness, True)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexNut
    Set ObjHexNut = Nothing
    Set centerPoint = Nothing
    Set oLineStr = Nothing
    
' Insert your code for output 9(davit Top Support 1) at the Lower side
    Dim topdavitSupLen As Double
    Dim topdavitSupTk As Double
    Dim ObjDavitTopSup1 As Object
    Dim davitCoordinate As Double
    Dim radialDist As Double
    Dim totalAngle As Double
        
    topdavitSupLen = davitPinDiameter + 0.05
    topdavitSupTk = 0.015
    davitCoordinate = Sqr(((flangeDiam) / 2) ^ 2 - ((topdavitSupLen / 2) ^ 2))
    angle = Atn((topdavitSupLen / 2) / davitCoordinate)
    totalAngle = davitRotationAngle - angle
    radialDist = (flangeDiam) / 2
    
    'Point 1
    dblPoints(0) = radialDist * Sin(totalAngle)
    dblPoints(1) = radialDist * Cos(totalAngle)
    dblPoints(2) = -(flangeThick) / 3 - topdavitSupTk
    'Point 2
    totalAngle = davitRotationAngle + angle
    dblPoints(3) = radialDist * Sin(totalAngle)
    dblPoints(4) = radialDist * Cos(totalAngle)
    dblPoints(5) = -(flangeThick) / 3 - topdavitSupTk
    'Point 3
    dblPoints(6) = radialDist * Sin(totalAngle)
    dblPoints(7) = radialDist * Cos(totalAngle)
    dblPoints(8) = -(flangeThick) / 3
    'Point4
    totalAngle = davitRotationAngle - angle
    dblPoints(9) = radialDist * Sin(totalAngle)
    dblPoints(10) = radialDist * Cos(totalAngle)
    dblPoints(11) = -(flangeThick) / 3
    'Point5
    dblPoints(12) = radialDist * Sin(totalAngle)
    dblPoints(13) = radialDist * Cos(totalAngle)
    dblPoints(14) = -(flangeThick) / 3 - topdavitSupTk
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set Sin(davitRotationAngle), Cos(davitRotationAngle), 0
    'davit Support
    Set ObjDavitTopSup1 = PlaceProjection(m_OutputColl, oLineStr, ProVect, topdavitSupWidth, True)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDavitTopSup1
    Set ObjDavitTopSup1 = Nothing
    Set oLineStr = Nothing
    
' Insert your code for output 10(davit Top Support 2)
    Dim topdavitSupLen1 As Double
    Dim topdavitSupWidth1 As Double
    
    topdavitSupLen1 = topdavitSupLen
    topdavitSupWidth1 = topdavitSupWidth + ((flangeDiam - pipeDiam)) / 2
    davitCoordinate = Sqr(((pipeDiam) / 2) ^ 2 - ((topdavitSupLen1 / 2) ^ 2))
    angle = Atn((topdavitSupLen1 / 2) / davitCoordinate)
    totalAngle = davitRotationAngle - angle
    radialDist = (pipeDiam) / 2
    
    'Point 1
    dblPoints(0) = radialDist * Sin(totalAngle)
    dblPoints(1) = radialDist * Cos(totalAngle)
    dblPoints(2) = -(flangeThick) / 3 - davitSupsSpacing
    'Point 2
    totalAngle = davitRotationAngle + angle
    dblPoints(3) = radialDist * Sin(totalAngle)
    dblPoints(4) = radialDist * Cos(totalAngle)
    dblPoints(5) = -(flangeThick) / 3 - davitSupsSpacing
    'Point 3
    dblPoints(6) = radialDist * Sin(totalAngle)
    dblPoints(7) = radialDist * Cos(totalAngle)
    dblPoints(8) = -(flangeThick) / 3 - davitSupsSpacing + topdavitSupTk1
    'Point4
    totalAngle = davitRotationAngle - angle
    dblPoints(9) = radialDist * Sin(totalAngle)
    dblPoints(10) = radialDist * Cos(totalAngle)
    dblPoints(11) = -(flangeThick) / 3 - davitSupsSpacing + topdavitSupTk1
    'Point5
    dblPoints(12) = radialDist * Sin(totalAngle)
    dblPoints(13) = radialDist * Cos(totalAngle)
    dblPoints(14) = -(flangeThick) / 3 - davitSupsSpacing
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set Sin(davitRotationAngle), Cos(davitRotationAngle), 0
    'davit Support
    Set ObjDavitTopSup1 = PlaceProjection(m_OutputColl, oLineStr, ProVect, topdavitSupWidth1, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDavitTopSup1
    Set ObjDavitTopSup1 = Nothing
    Set oLineStr = Nothing
    
' Insert your code for output 11(Davit Vertical Support)
    Dim ObjDavitVertSup As Object
    
    davitCoordinate = Sqr((flangeDiam / 2) ^ 2 - (topdavitSupLen / 2) ^ 2) + topdavitSupWidth _
                                    - topdavitSupTk1
    angle = Atn((topdavitSupLen / 2) / (davitCoordinate))
    radialDist = (topdavitSupLen / 2) / Sin(angle)
    totalAngle = davitRotationAngle - angle
    
    'Point 1
    dblPoints(0) = radialDist * Sin(totalAngle)
    dblPoints(1) = radialDist * Cos(totalAngle)
    dblPoints(2) = -flangeThick / 3
    'Point 2
    totalAngle = davitRotationAngle + angle
    dblPoints(3) = radialDist * Sin(totalAngle)
    dblPoints(4) = radialDist * Cos(totalAngle)
    dblPoints(5) = -flangeThick / 3
    'Point 3
    dblPoints(6) = radialDist * Sin(totalAngle)
    dblPoints(7) = radialDist * Cos(totalAngle)
    dblPoints(8) = -flangeThick / 3 - davitSupsSpacing
    'Point4
    totalAngle = davitRotationAngle - angle
    dblPoints(9) = radialDist * Sin(totalAngle)
    dblPoints(10) = radialDist * Cos(totalAngle)
    dblPoints(11) = -flangeThick / 3 - davitSupsSpacing
    'Point5
    dblPoints(12) = radialDist * Sin(totalAngle)
    dblPoints(13) = radialDist * Cos(totalAngle)
    dblPoints(14) = -flangeThick / 3
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set Sin(davitRotationAngle), Cos(davitRotationAngle), 0
    'davit Support
    Set ObjDavitVertSup = PlaceProjection(m_OutputColl, oLineStr, ProVect, topdavitSupTk1, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDavitVertSup
    Set ObjDavitVertSup = Nothing
    Set oLineStr = Nothing
    Set ProVect = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
