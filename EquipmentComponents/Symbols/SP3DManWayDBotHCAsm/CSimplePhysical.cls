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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:33 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, May 13 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Davit for Bottom Horizontal Cover, is prepared as per DOW Chemical Company manual.
'   Dimensions are considered for two pipe size viz., 20" and 24" . All assumed values are as per the DOW
'   documents unless otherwise specified.
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
    Dim NozzleLength As Double
    Dim parInsulationThickness As Double
    Dim parCentertoDavitCover As Double
    Dim parDomeHeight1 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parDomeHeight1 = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
           
    'The value "davitRotationAngle" can be changed to rotate the Davit about the nozzle axis. As the angle is measured
    'from North and as per the DOW document the davit is aligned along East Davit Rotation angle of 90Degree
    'is considered for this Symbol.
    
    Dim davitRotationAngle As Double
    davitRotationAngle = PI / 2

' Place Nozzle 1
    Dim ObjManhole             As IJDNozzle
    Dim oDir As AutoMath.DVector
    Dim oPlacePoint As AutoMath.DPosition
    

    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    

    'Nozzle length calculation: Nozzle is taken inside of vessel by 20% of dome height (parDomeHeight1)
    NozzleLength = parCentertoDavitCover - parDomeHeight1 + 0.2 * parDomeHeight1
    'If the nozzle length is than 0.1, set the nozzle length equal to flange thickness
    If NozzleLength <= 0.1 Then
        NozzleLength = flangeThick
    End If

    oDir.Set 0, 0, -1
    oPlacePoint.Set 0, 0, 0

    Set ObjManhole = CreateNozzlePHWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjManhole
    Set ObjManhole = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
' Insert your code for output 2(Blind Flange)
    Dim flangeGap As Double
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    flangeGap = 0.004           'Gap between flanges is assumed as 4mm
    stPoint.Set 0, 0, -flangeGap
    enPoint.Set 0, 0, -flangeThick - flangeGap
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, flangeDiam

' Insert your code for output 3(BottomSupport Plate)
    'Create a line string and project it
    Dim plateThickness As Double '(Dimension along X-axis)
    Dim plateDepth1 As Double 'Minimum thickness
    Dim plateDepth2 As Double 'Maximum Thickness
    Dim plateSlopeLen As Double
    Dim plateLen As Double
    Dim dblPoints(0 To 17) As Double
    Dim ProVect As AutoMath.DVector
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim ObjBotSupPlate As Object
    Dim radialLength As Double
    Dim angle As Double
    Dim totalAngle As Double
    Dim davitSupDepth As Double
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set ProVect = New AutoMath.DVector
    
    'Values as per DOW document
    davitSupDepth = 0.178
    plateThickness = 0.012
    plateDepth1 = 0.025
    plateDepth2 = flangeGap + (flangeThick) / 3 + davitSupDepth - _
                            (flangeGap + flangeThick)
    plateLen = flangeDiam / 2
    plateSlopeLen = plateLen * 0.8
    
    'Point 1
    dblPoints(0) = -(plateThickness / 2) * Cos(davitRotationAngle)
    dblPoints(1) = (plateThickness / 2) * Sin(davitRotationAngle)
    dblPoints(2) = -flangeThick - flangeGap
    
    'Point 2
    dblPoints(3) = -(plateThickness / 2) * Cos(davitRotationAngle)
    dblPoints(4) = (plateThickness / 2) * Sin(davitRotationAngle)
    dblPoints(5) = -flangeThick - flangeGap - plateDepth1
    
    'Point 3
    radialLength = Sqr((plateThickness / 2) ^ 2 + plateSlopeLen ^ 2)
    angle = Atn((plateThickness / 2) / plateSlopeLen)
    totalAngle = davitRotationAngle - angle
    
    dblPoints(6) = radialLength * Sin(totalAngle)
    dblPoints(7) = radialLength * Cos(totalAngle)
    dblPoints(8) = -flangeThick - flangeGap - plateDepth2
    
    'Point4
    radialLength = flangeDiam / 2
    plateLen = Sqr(radialLength ^ 2 - (plateThickness / 2) ^ 2)
    angle = Atn((plateThickness / 2) / plateLen)
    totalAngle = davitRotationAngle - angle
    
    dblPoints(9) = radialLength * Sin(totalAngle)
    dblPoints(10) = radialLength * Cos(totalAngle)
    dblPoints(11) = -flangeThick - flangeGap - plateDepth2
    
    'Point5
    dblPoints(12) = radialLength * Sin(totalAngle)
    dblPoints(13) = radialLength * Cos(totalAngle)
    dblPoints(14) = -flangeThick - flangeGap
    
    'Point6
    dblPoints(15) = dblPoints(0)
    dblPoints(16) = dblPoints(1)
    dblPoints(17) = dblPoints(2)
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dblPoints)
    ProVect.Set Cos(davitRotationAngle), -Sin(davitRotationAngle), 0
    'Support
    Set ObjBotSupPlate = PlaceProjection(m_OutputColl, oLineStr, ProVect, plateThickness, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotSupPlate
    Set ObjBotSupPlate = Nothing
    
' Insert your code for output 4(Bottom davit Support)
'To create a box draw a line string and project it along Y-axis
    Dim davitSupWidth As Double     '(Along Y-axis)
    Dim davitSupLength As Double    '(Along X-axis)
    Dim pinDist As Double
    Dim davitCoordinate As Double
    Dim ObjBotDavitSup As Object
    

    pinDist = 0.064
    davitSupWidth = 2 * pinDist     'Assumed to be twice pinDist
    davitSupLength = 0.15
    davitCoordinate = Sqr((flangeDiam / 2) ^ 2 - (davitSupLength / 2) ^ 2)
    angle = Atn((davitSupLength / 2) / davitCoordinate)
    totalAngle = davitRotationAngle - angle
    plateLen = flangeDiam / 2
    
    'Point 1
    dblPoints(0) = plateLen * Sin(totalAngle)
    dblPoints(1) = plateLen * Cos(totalAngle)
    dblPoints(2) = -flangeGap - (flangeThick) / 3
    
    'Point 2
    totalAngle = davitRotationAngle + angle     'Point is in the other side of X-axis
    
    dblPoints(3) = plateLen * Sin(totalAngle)
    dblPoints(4) = plateLen * Cos(totalAngle)
    dblPoints(5) = -flangeGap - (flangeThick) / 3
    'Point 3
    dblPoints(6) = plateLen * Sin(totalAngle)
    dblPoints(7) = plateLen * Cos(totalAngle)
    dblPoints(8) = -flangeGap - (flangeThick) / 3 - davitSupDepth
    
    'Point4
    totalAngle = davitRotationAngle - angle
    dblPoints(9) = plateLen * Sin(totalAngle)
    dblPoints(10) = plateLen * Cos(totalAngle)
    dblPoints(11) = -flangeGap - (flangeThick) / 3 - davitSupDepth
    'Point5
    dblPoints(12) = dblPoints(0)
    dblPoints(13) = dblPoints(1)
    dblPoints(14) = dblPoints(2)
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set Sin(davitRotationAngle), Cos(davitRotationAngle), 0
    
    'davit Support
    Set ObjBotDavitSup = PlaceProjection(m_OutputColl, oLineStr, ProVect, davitSupWidth, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotDavitSup
    Set ObjBotDavitSup = Nothing
    
' Insert your code for output 5(davit Top Support 1) at the Lower side
    Dim topdavitSupLen As Double
    Dim topdavitSupWidth As Double
    Dim topdavitSupTk As Double
    Dim ObjTopDavitSup As Object
    Dim pinDiameter As Double
    
    pinDiameter = 0.048
    
    topdavitSupLen = 0.15
    topdavitSupWidth = pinDist + pinDiameter
    topdavitSupTk = 0.012
    davitCoordinate = Sqr((flangeDiam / 2) ^ 2 - (topdavitSupLen / 2) ^ 2)
    angle = Atn((topdavitSupLen / 2) / davitCoordinate)
    totalAngle = davitRotationAngle - angle
    
    'Point 1
    dblPoints(0) = plateLen * Sin(totalAngle)
    dblPoints(1) = plateLen * Cos(totalAngle)
    dblPoints(2) = flangeThick / 3 + topdavitSupTk
    'Point 2
    totalAngle = davitRotationAngle + angle
    
    dblPoints(3) = plateLen * Sin(totalAngle)
    dblPoints(4) = plateLen * Cos(totalAngle)
    dblPoints(5) = flangeThick / 3 + topdavitSupTk
    'Point 3
    dblPoints(6) = plateLen * Sin(totalAngle)
    dblPoints(7) = plateLen * Cos(totalAngle)
    dblPoints(8) = flangeThick / 3
    'Point4
    totalAngle = davitRotationAngle - angle
    
    dblPoints(9) = plateLen * Sin(totalAngle)
    dblPoints(10) = plateLen * Cos(totalAngle)
    dblPoints(11) = flangeThick / 3
    'Point5
    dblPoints(12) = dblPoints(0)
    dblPoints(13) = dblPoints(1)
    dblPoints(14) = dblPoints(2)
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set Sin(davitRotationAngle), Cos(davitRotationAngle), 0
    
    'davit Support
    Set ObjTopDavitSup = PlaceProjection(m_OutputColl, oLineStr, ProVect, topdavitSupWidth, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopDavitSup
    Set ObjTopDavitSup = Nothing
    
' Insert your code for output 6(davit Top Support 2)
    Dim topdavitSupLen1 As Double
    Dim topdavitSupWidth1 As Double
    Dim topdavitSupTk1 As Double
    Dim davitSupsSpacing As Double
    
    topdavitSupLen1 = 0.15
    topdavitSupWidth1 = topdavitSupWidth + ((flangeDiam - pipeDiam)) / 2
    topdavitSupTk1 = 0.012
    davitSupsSpacing = 0.15
    davitCoordinate = Sqr((pipeDiam / 2) ^ 2 - (topdavitSupLen1 / 2) ^ 2)
    angle = Atn((topdavitSupLen1 / 2) / davitCoordinate)
    totalAngle = davitRotationAngle - angle
    plateLen = 0.5 * pipeDiam
    
    'Point 1
    dblPoints(0) = plateLen * Sin(totalAngle)
    dblPoints(1) = plateLen * Cos(totalAngle)
    dblPoints(2) = flangeThick / 3 + davitSupsSpacing
    'Point 2
    totalAngle = davitRotationAngle + angle
    dblPoints(3) = plateLen * Sin(totalAngle)
    dblPoints(4) = plateLen * Cos(totalAngle)
    dblPoints(5) = flangeThick / 3 + davitSupsSpacing
    'Point 3
    dblPoints(6) = plateLen * Sin(totalAngle)
    dblPoints(7) = plateLen * Cos(totalAngle)
    dblPoints(8) = flangeThick / 3 + davitSupsSpacing - topdavitSupTk1
    'Point4
    totalAngle = davitRotationAngle - angle
    dblPoints(9) = plateLen * Sin(totalAngle)
    dblPoints(10) = plateLen * Cos(totalAngle)
    dblPoints(11) = flangeThick / 3 + davitSupsSpacing - topdavitSupTk1
    'Point5
    dblPoints(12) = dblPoints(0)
    dblPoints(13) = dblPoints(1)
    dblPoints(14) = dblPoints(2)
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPoints)
    ProVect.Set Sin(davitRotationAngle), Cos(davitRotationAngle), 0
    'davit Support
    Set ObjTopDavitSup = PlaceProjection(m_OutputColl, oLineStr, ProVect, topdavitSupWidth1, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopDavitSup
    Set ObjTopDavitSup = Nothing
    
' Insert your code for output 7(davit Pin)
    Dim pinExtn As Double
    
    'Pin extension after the davit bottom support is 0.5" + 1"+3*pinDiameter (3*D)
    pinExtn = 0.012 + 0.025 + 3 * pinDiameter
    stPoint.Set (flangeDiam / 2 + pinDist) * Sin(davitRotationAngle), _
                        (flangeDiam / 2 + pinDist) * Cos(davitRotationAngle), _
                    -flangeGap - flangeThick / 3 - davitSupDepth - pinExtn
    enPoint.Set stPoint.x, stPoint.y, _
                                    flangeThick / 3 + davitSupsSpacing + pinDiameter
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, pinDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 8(Hexagonal Nut)
    Dim LineStrPoints(0 To 20) As Double
    Dim centerPoint As AutoMath.DPosition
    Dim hexNutDia As Double
    Dim nutThickness As Double
    Dim i As Integer
    Dim washerTk As Double
    Dim ObjHexNut As Object
    Set centerPoint = New AutoMath.DPosition
    
    hexNutDia = pinDiameter * 1.4
    nutThickness = 0.025    'As per DOW Document
    washerTk = 0.012
    centerPoint.Set (flangeDiam / 2 + pinDist) * Sin(davitRotationAngle), _
                                (flangeDiam / 2 + pinDist) * Cos(davitRotationAngle), _
                                -(flangeThick) / 3 - davitSupDepth - washerTk
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
    ProVect.Set 0, 0, -1
    
    Set ObjHexNut = PlaceProjection(m_OutputColl, oLineStr, ProVect, nutThickness, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexNut
    Set ObjHexNut = Nothing
    Set centerPoint = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
