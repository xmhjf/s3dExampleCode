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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Monday, Mar 25 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   SP3DBioClamp Symbol is created with Seven number of out puts.
'   Out put: 1 - Line string is formed and rotated about X-axis. Out put:2 - Upper Screw support of Box
'   Out put: 3 - Lower Screw Support of Box.Out put: 4 - Stem of Screw using Cylinder. Out put: 5 Handle
'   using Cuylinder and 2 number of Nozzle
'
'                               |------------|         ---
'                               |--------|   |
'                                        |   |          Clamp Width
' Sidewall                      |--------|   |
' thickness  =>                 |------------|         ---
'
'                          Clamp Total Tk|   |
'   The above cross section will be rotated about X-Axis
'
'   Change History:
'   dd.mmm.yyyy      who                 change description
'   ---------        -----               ------------------
'   09.Jul.2003     SymbolTeam(India)    Copyright Information, Header  is added.
'   10.Sep.2004     SymbolTeam(India)    CR 28365 Removed nozzle creation code.
'   01.Jan.2005     svsmylav             TR-70434: Defined InsulationThickness input.
'   08.SEP.2006     KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'   12.Oct.2007     VRG                  CR-127155 Enhanced the existing  BioClamp to support all Tri-Clamp options using PartDataBasis values.
'   15.Nov.2007     VRG                  CR-127647 Enhanced the symbol code to support Hinged Clamp for joining Kwik Flanges
'   15.Nov.2007     VRG/PK               CR-127156 Enhanced the symbol code to support Swivel Joint Clamp
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:"    'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim iOutput As Integer
    Dim iCount As Integer
    Dim ObjClampBody As Object
    Dim ObjUpperSup As Object
    Dim ObjLowSup As Object
    Dim ObjStem As Object
    Dim ObjHandle As Object
    Dim oProjection As Object

    Dim oBolt As Object
    Dim oLine As Object
    Dim lines As Collection
    Dim parClampDiameter As Double
    Dim parClampLength As Double
    Dim parClampWidth As Double
    Dim parInsulationThickness As Double
    Dim ClampTk As Double, ClampWidth As Double, ClampSideWallTk As Double, ClampTotalTk As Double
    Dim ClampOpening As Double, InSideRadius As Double, OutSideRadius As Double, includAngle As Double
    Dim HalfOpeningAngle As Double, DistalongAxis As Double, ScrewSupLen As Double
    Dim StemTotalHt As Double, StemDia As Double, ScrewHandleLen As Double, HandleDia As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampDiameter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    If UBound(arrayOfInputs) >= 4 Then     'parClampLength and parClampWidth are used for Hinged Clamp
        parClampLength = arrayOfInputs(4)
        parClampWidth = arrayOfInputs(5)
    End If
    'This symbol can be associated with the following Part data basis values

    '1.Clamp, double hinge heavy duty with wing nut - 430 (The default case)
    '2.Clamp, single hinge heavy duty with wing nut - 431
    '3.Clamp, 3-segment heavy duty with wing nut - 432
    '4.Clamp, medium duty with wing nut - 433
    '5.Clamp, light duty with t-bolt - 434
    '6.Clamp, light duty adjustable - 435
    '7.Clamp, fractional with wing nut - 436
    '8.Clamp, high pressure -437
    '9.Clamp, sanitary for hose adapter - 438
    '10.Clamp, band - 439
    '11.Clamp, hinged - 440
    '12.Clamp, swivel joint - 441

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oAxis As New AutoMath.DVector
    Dim oArc As IngrGeom3D.Arc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oTransMatrix As IJDT4x4
    If (lPartDataBasis = 439) Then
        parClampDiameter = parClampDiameter / 1.1
    End If
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    If (lPartDataBasis = 440) Then
        parClampDiameter = parClampDiameter / 1.1
        ClampSideWallTk = flangeThick
        ClampWidth = parClampWidth
        InSideRadius = 0.45 * flangeDiam
        OutSideRadius = parClampDiameter * 1.1 / 2
        ClampTotalTk = 0.5 * (parClampLength - parClampDiameter * 1.1)
        ClampOpening = parClampDiameter * 0.05
        'Angle Subtended  between the opening is determined using Phythocross Theorm.
        DistalongAxis = Sqr((InSideRadius * InSideRadius) - ((ClampOpening * 0.5) * (ClampOpening * 0.5)))
        HalfOpeningAngle = Atn((ClampOpening * 0.5) / DistalongAxis)
        includAngle = HalfOpeningAngle * 2
        StemTotalHt = flangeThick * 4 + ClampTotalTk * 2 + ClampOpening
        ScrewSupLen = Cos(HalfOpeningAngle) * (OutSideRadius - InSideRadius) + 0.7 * (parClampLength - parClampDiameter * 1.1)
        StemDia = flangeThick
        HandleDia = flangeThick
        ScrewHandleLen = flangeThick * 4
    Else
        'Clamp thickness is Half of Difference between ((parClampDiameter)-(1.1*parClampDiameter))
        ClampTk = ((1.1 * parClampDiameter) - (parClampDiameter)) / 2
        ClampSideWallTk = flangeThick
        'Clamp Inside Width is 1.01 times Twice the Flange Thickness. For Total Clamp width add side wall
        'thickness
        ClampWidth = 1.01 * 2 * flangeThick + ClampSideWallTk + ClampSideWallTk
        InSideRadius = 0.9 * parClampDiameter / 2
        Dim x As Double
        If (lPartDataBasis = 431 Or lPartDataBasis = 434 Or lPartDataBasis = 436) Then   'Partdatabasis value 431 is used for Tri-Clover A13MHM clamp
            x = 1.1
        ElseIf (lPartDataBasis = 437) Then
            x = 1.2
        Else
            x = 1
        End If
        OutSideRadius = x * 1.1 * parClampDiameter / 2
        ClampTotalTk = x * ((1.1 * parClampDiameter) - (0.9 * parClampDiameter)) / 2
    
        'Distance between Axis center and Clamp Arc starting point  is taken as  0.01*parClampDiameter
        'or 2mm Which ever is greater
        ClampOpening = parClampDiameter * 0.01 * 2
        If ClampOpening < (0.002 * 2) Then ClampOpening = (0.002 * 2)
        If (lPartDataBasis = 437 Or lPartDataBasis = 438 Or lPartDataBasis = 439) Then
            ClampOpening = 0
        ElseIf (lPartDataBasis = 434) Then
            ClampOpening = parClampDiameter * 0.2
        ElseIf (lPartDataBasis = 435) Then
            ClampOpening = parClampDiameter * 0.4
        End If
        'Angle Subtended  between the opening is determined using Phythocross Theorm.
        DistalongAxis = Sqr((InSideRadius * InSideRadius) - ((ClampOpening * 0.5) * (ClampOpening * 0.5)))
        HalfOpeningAngle = Atn((ClampOpening * 0.5) / DistalongAxis)
        includAngle = HalfOpeningAngle * 2
    
        'Stem Height is 4 Times of Flange thickness from the top of Clamp opening. To get height from the
        'bottom of Clamp opening Add Bottom and Top thickness of Clamp Screw support and clamp opening
        StemTotalHt = flangeThick * 4 + ClampTotalTk * 2 + ClampOpening
        ScrewSupLen = Cos(HalfOpeningAngle) * (OutSideRadius - InSideRadius) + flangeThick * 4
        StemDia = flangeThick
        HandleDia = flangeThick
        ScrewHandleLen = flangeThick * 4
    End If

    iOutput = 0

    ' Insert your code for output 1(ClampBody)
    Dim RevVector As AutoMath.DVector
    Dim CenPtDPosition As AutoMath.DPosition
    Dim revAngle As Double

    Set RevVector = New AutoMath.DVector
    Set CenPtDPosition = New AutoMath.DPosition
    revAngle = 2 * PI - includAngle
    If (lPartDataBasis = 433) Then
        revAngle = PI - includAngle
    ElseIf (lPartDataBasis = 441) Then
        revAngle = 2 * PI
    End If
    If (lPartDataBasis = 441) Then
        Dim dLineStrPts(0 To 38) As Double
        'Point at Negative  Y-axis and X-axis near Center(1point)
        dLineStrPts(0) = -ClampWidth / 2                                              'X-axis Value
        dLineStrPts(1) = -InSideRadius * Cos(HalfOpeningAngle)                        'Y-Axis Value
        dLineStrPts(2) = ClampOpening / 2                                             'Z-Axis Value
    
        'Point at Negative  Y-axis and X-axis away from Center (2point)
        dLineStrPts(3) = -ClampWidth / 2                                              'X-axis Value
        dLineStrPts(4) = -OutSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        dLineStrPts(5) = OutSideRadius * Sin(HalfOpeningAngle)                        'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis away from Center(3point)
        dLineStrPts(6) = ClampWidth / 2                                               'X-axis Value
        dLineStrPts(7) = -OutSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        dLineStrPts(8) = OutSideRadius * Sin(HalfOpeningAngle)                        'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis near Center (4point)
        dLineStrPts(9) = ClampWidth / 2                                               'X-axis Value
        dLineStrPts(10) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        dLineStrPts(11) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis near Center (5point)
        dLineStrPts(12) = ClampWidth / 2 - 0.7 * ClampSideWallTk                      'X-axis Value
        dLineStrPts(13) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        dLineStrPts(14) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis inner wall of clamp (6point)
        dLineStrPts(15) = ClampWidth / 2 - 0.7 * ClampSideWallTk                      'X-axis Value
        dLineStrPts(16) = -0.9 * InSideRadius * Cos(HalfOpeningAngle)                 'Y-Axis Value
        dLineStrPts(17) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis inner wall of clamp (7point)
        dLineStrPts(18) = ClampWidth / 2 - 1.2 * ClampSideWallTk                      'X-axis Value
        dLineStrPts(19) = -0.9 * InSideRadius * Cos(HalfOpeningAngle)                 'Y-Axis Value
        dLineStrPts(20) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative Y-axis and Negative X-axis inner wall of clamp (8point)
        dLineStrPts(21) = ClampWidth / 2 - 1.4 * ClampSideWallTk                      'X-axis Value
        dLineStrPts(22) = -(0.47 * parClampDiameter) * Cos(HalfOpeningAngle)          'Y-Axis Value
        dLineStrPts(23) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative Y-axis and Negative X-axis inner wall of clamp (9point)
        dLineStrPts(24) = -ClampWidth / 2 + 1.4 * ClampSideWallTk                     'X-axis Value
        dLineStrPts(25) = -(0.47 * parClampDiameter) * Cos(HalfOpeningAngle)          'Y-Axis Value
        dLineStrPts(26) = ClampOpening / 2                                            'Z-Axis Value
        
        'Point at Negative  Y-axis and Positive X-axis inner wall of clamp (10point)
        dLineStrPts(27) = -ClampWidth / 2 + 1.2 * ClampSideWallTk                     'X-axis Value
        dLineStrPts(28) = -0.9 * InSideRadius * Cos(HalfOpeningAngle)                 'Y-Axis Value
        dLineStrPts(29) = ClampOpening / 2                                            'Z-Axis Value
    
       'Point at Negative  Y-axis and Positive X-axis inner wall of clamp (11point)
        dLineStrPts(30) = -ClampWidth / 2 + 0.7 * ClampSideWallTk                     'X-axis Value
        dLineStrPts(31) = -0.9 * InSideRadius * Cos(HalfOpeningAngle)                 'Y-Axis Value
        dLineStrPts(32) = ClampOpening / 2                                            'Z-Axis Value
        
        'Point at Negative Y-axis and Negative X-axis (12point)
        dLineStrPts(33) = -ClampWidth / 2 + 0.7 * ClampSideWallTk                     'X-axis Value
        dLineStrPts(34) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        dLineStrPts(35) = ClampOpening / 2                                            'Z-Axis Value
    
        'Starting Point
        dLineStrPts(36) = -ClampWidth / 2                                             'X-axis Value
        dLineStrPts(37) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        dLineStrPts(38) = ClampOpening / 2                                            'Z-Axis Value
    
        RevVector.Set -1, 0, 0
        CenPtDPosition.Set 0, 0, 0
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 13, dLineStrPts)
        
    Else
        Dim LineStrPoints(0 To 26) As Double
        'Point at Negative  Y-axis and X-axis near Center(1point)
        LineStrPoints(0) = -ClampWidth / 2                                              'X-axis Value
        LineStrPoints(1) = -InSideRadius * Cos(HalfOpeningAngle)                        'Y-Axis Value
        LineStrPoints(2) = ClampOpening / 2                                             'Z-Axis Value
    
        'Point at Negative  Y-axis and X-axis away from Center (2point)
        LineStrPoints(3) = -ClampWidth / 2                                              'X-axis Value
        LineStrPoints(4) = -OutSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        LineStrPoints(5) = OutSideRadius * Sin(HalfOpeningAngle)                        'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis away from Center(3point)
        LineStrPoints(6) = ClampWidth / 2                                               'X-axis Value
        LineStrPoints(7) = -OutSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        LineStrPoints(8) = OutSideRadius * Sin(HalfOpeningAngle)                        'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis near Center (4point)
        LineStrPoints(9) = ClampWidth / 2                                               'X-axis Value
        LineStrPoints(10) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        LineStrPoints(11) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis near Center (5point)
        LineStrPoints(12) = ClampWidth / 2 - ClampSideWallTk                            'X-axis Value
        LineStrPoints(13) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        LineStrPoints(14) = ClampOpening / 2                                            'Z-Axis Value
    
        'Point at Negative  Y-axis and Positive X-axis inner wall of clamp (6point)
        LineStrPoints(15) = ClampWidth / 2 - ClampSideWallTk                            'X-axis Value
        LineStrPoints(16) = -(parClampDiameter / 2) * Cos(HalfOpeningAngle)             'Y-Axis Value
        LineStrPoints(17) = (parClampDiameter / 2) * Sin(HalfOpeningAngle)              'Z-Axis Value
    
        'Point at Negative Y-axis and Negative X-axis inner wall of clamp (7point)
        LineStrPoints(18) = -ClampWidth / 2 + ClampSideWallTk                           'X-axis Value
        LineStrPoints(19) = -(parClampDiameter / 2) * Cos(HalfOpeningAngle)             'Y-Axis Value
        LineStrPoints(20) = (parClampDiameter / 2) * Sin(HalfOpeningAngle)              'Z-Axis Value
    
        'Point at Negative Y-axis and Negative X-axis (8point)
        LineStrPoints(21) = -ClampWidth / 2 + ClampSideWallTk                           'X-axis Value
        LineStrPoints(22) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        LineStrPoints(23) = ClampOpening / 2                                            'Z-Axis Value
    
        'Starting Point
        LineStrPoints(24) = -ClampWidth / 2                                             'X-axis Value
        LineStrPoints(25) = -InSideRadius * Cos(HalfOpeningAngle)                       'Y-Axis Value
        LineStrPoints(26) = ClampOpening / 2                                            'Z-Axis Value
    
        RevVector.Set -1, 0, 0
        CenPtDPosition.Set 0, 0, 0
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStrPoints)
    End If
    Set ObjClampBody = PlaceRevolution(m_OutputColl, oLineString, RevVector, CenPtDPosition, revAngle, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampBody
    Set ObjClampBody = Nothing
    If (lPartDataBasis = 433) Then
        Set ObjClampBody = PlaceRevolution(m_OutputColl, oLineString, RevVector, CenPtDPosition, revAngle, True)
        Set oTransMatrix = New AutoMath.DT4x4
        oTransMatrix.LoadIdentity
        oTransMatrix.IndexValue(10) = -1
        ObjClampBody.Transform oTransMatrix
        m_OutputColl.AddOutput "ClampSegment", ObjClampBody
    End If
    Set CenPtDPosition = Nothing
    Set RevVector = Nothing
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    objLineString.Remove

    'Insert your code for output 2(upperSup)
    If (lPartDataBasis = 435) Then
        Set RevVector = New AutoMath.DVector
        Set CenPtDPosition = New AutoMath.DPosition
        revAngle = PI
        RevVector.Set -1, 0, 0
        CenPtDPosition.Set 0, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, 0.8 * ClampOpening
        Dim oHook As Object
        Set oHook = PlaceRevolution(m_OutputColl, oLineString, RevVector, CenPtDPosition, revAngle, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHook
        Set oHook = Nothing
        Set oHook = PlaceRevolution(m_OutputColl, oLineString, RevVector, CenPtDPosition, revAngle, True)
        Set oTransMatrix = New AutoMath.DT4x4
        oTransMatrix.LoadIdentity
        oTransMatrix.IndexValue(10) = -1
        oHook.Transform oTransMatrix
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHook
        Set oHook = Nothing
        Set RevVector = Nothing
        Set CenPtDPosition = Nothing
    Else
        Dim BoxStrPoints(0 To 14) As Double
        Dim ProjVect As AutoMath.DVector
        Set ProjVect = New AutoMath.DVector
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        If (lPartDataBasis = 434 Or lPartDataBasis = 438 Or lPartDataBasis = 439) Then
            oStPoint.Set -ClampWidth / 2, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, OutSideRadius * Sin(PI / 9)
            oEnPoint.Set ClampWidth / 2, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, OutSideRadius * Sin(PI / 9)
            Set ObjUpperSup = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * flangeThick, True)
        ElseIf (lPartDataBasis = 437) Then
            oStPoint.Set -0.4 * ClampWidth, 0.4 * parClampDiameter, 0.5 * parClampDiameter
            oEnPoint.Set 0.4 * ClampWidth, -0.4 * parClampDiameter, 0.8 * parClampDiameter
            Set ObjUpperSup = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        Else
            'Bottom Left Point
            BoxStrPoints(0) = -ClampWidth / 2
            BoxStrPoints(1) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(2) = ClampOpening / 2
    
            'Bottom Right Point
            BoxStrPoints(3) = ClampWidth / 2
            BoxStrPoints(4) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(5) = ClampOpening / 2
    
            'Top Right Point
            BoxStrPoints(6) = ClampWidth / 2
            BoxStrPoints(7) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(8) = ClampOpening / 2 + ClampTotalTk
    
            'Top Left Point
            BoxStrPoints(9) = -ClampWidth / 2
            BoxStrPoints(10) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(11) = ClampOpening / 2 + ClampTotalTk
    
            'Starting Point
            BoxStrPoints(12) = -ClampWidth / 2
            BoxStrPoints(13) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(14) = ClampOpening / 2
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, BoxStrPoints)
    
            ProjVect.Set 0, -1, 0
            Set ObjUpperSup = PlaceProjection(m_OutputColl, oLineString, ProjVect, ScrewSupLen, True)
            Set ProjVect = Nothing
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
        End If
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperSup
        Set ObjUpperSup = Nothing
    End If
    'Insert your code for output 3(LowSup)
    If (lPartDataBasis = 435) Then
    
    Else
        Dim ProjLowVect As AutoMath.DVector
        Set ProjLowVect = New AutoMath.DVector
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        If (lPartDataBasis = 434 Or lPartDataBasis = 438 Or lPartDataBasis = 439) Then
            oStPoint.Set -ClampWidth / 2, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 4)
            oEnPoint.Set ClampWidth / 2, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 4)
            If (lPartDataBasis = 439 Or lPartDataBasis = 438) Then
                oStPoint.Set -ClampWidth / 2, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 9)
                oEnPoint.Set ClampWidth / 2, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 9)
            End If
            Set ObjLowSup = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * flangeThick, True)
        ElseIf (lPartDataBasis = 437) Then
            oStPoint.Set -0.4 * ClampWidth, 0.4 * parClampDiameter, -0.5 * parClampDiameter
            oEnPoint.Set 0.4 * ClampWidth, -0.4 * parClampDiameter, -0.8 * parClampDiameter
            Set ObjLowSup = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        Else
            'Top Left Point
            BoxStrPoints(0) = -ClampWidth / 2
            BoxStrPoints(1) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(2) = -ClampOpening / 2
    
            'Top Right Point
            BoxStrPoints(3) = ClampWidth / 2
            BoxStrPoints(4) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(5) = -ClampOpening / 2
    
            'Bottom Right Point
            BoxStrPoints(6) = ClampWidth / 2
            BoxStrPoints(7) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(8) = -ClampOpening / 2 - ClampTotalTk
    
            'Bottom Left Point
            BoxStrPoints(9) = -ClampWidth / 2
            BoxStrPoints(10) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(11) = -ClampOpening / 2 - ClampTotalTk
    
            'Starting Point
            BoxStrPoints(12) = -ClampWidth / 2
            BoxStrPoints(13) = -InSideRadius * Cos(HalfOpeningAngle)
            BoxStrPoints(14) = -ClampOpening / 2
    
            ProjLowVect.Set 0, -1, 0
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, BoxStrPoints)
            Set ObjLowSup = PlaceProjection(m_OutputColl, oLineString, ProjLowVect, ScrewSupLen, True)
            Set ProjLowVect = Nothing
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
        End If
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLowSup
        Set ObjLowSup = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    End If
    'Insert your code for output (Hinge)
    If (lPartDataBasis = 431 Or lPartDataBasis = 436 Or lPartDataBasis = 432 Or lPartDataBasis = 433 Or lPartDataBasis = 440) Then
        Dim oHinge As Object
        Dim HingeStPt As AutoMath.DPosition
        Dim HingeEndPt As AutoMath.DPosition
        Set HingeStPt = New AutoMath.DPosition
        Set HingeEndPt = New AutoMath.DPosition
        HingeStPt.Set -ClampWidth / 2, OutSideRadius, 0
        HingeEndPt.Set ClampWidth / 2, OutSideRadius, 0
        If (lPartDataBasis = 432) Then
            HingeStPt.Set -ClampWidth / 2, OutSideRadius * Cos(PI / 4), OutSideRadius * Sin(PI / 4)
            HingeEndPt.Set ClampWidth / 2, OutSideRadius * Cos(PI / 4), OutSideRadius * Sin(PI / 4)
        End If
        If (lPartDataBasis = 433) Then
            HingeStPt.Set -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle), OutSideRadius * Sin(HalfOpeningAngle) + 0.5 * flangeThick
            HingeEndPt.Set ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle), OutSideRadius * Sin(HalfOpeningAngle) + 0.5 * flangeThick
            Set oHinge = PlaceCylinder(m_OutputColl, HingeStPt, HingeEndPt, flangeThick, True)
        ElseIf (lPartDataBasis = 440) Then
            Set oHinge = PlaceCylinder(m_OutputColl, HingeStPt, HingeEndPt, 0.6 * (parClampLength - parClampDiameter * 1.1), True)
        Else
            Set oHinge = PlaceCylinder(m_OutputColl, HingeStPt, HingeEndPt, 2 * flangeThick, True)
        End If
        'Set the output
        m_OutputColl.AddOutput "Hinge_", oHinge
        Set oHinge = Nothing
        If (lPartDataBasis = 432 Or lPartDataBasis = 433) Then
            HingeStPt.Set -ClampWidth / 2, OutSideRadius * Cos(PI / 4), -OutSideRadius * Sin(PI / 4)
            HingeEndPt.Set ClampWidth / 2, OutSideRadius * Cos(PI / 4), -OutSideRadius * Sin(PI / 4)
            If (lPartDataBasis = 433) Then
                HingeStPt.Set -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle), -OutSideRadius * Sin(HalfOpeningAngle) - 0.5 * flangeThick
                HingeEndPt.Set ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle), -OutSideRadius * Sin(HalfOpeningAngle) - 0.5 * flangeThick
                Set oHinge = PlaceCylinder(m_OutputColl, HingeStPt, HingeEndPt, flangeThick, True)
            Else
                Set oHinge = PlaceCylinder(m_OutputColl, HingeStPt, HingeEndPt, 2 * flangeThick, True)
            End If
            m_OutputColl.AddOutput "Hinge_", oHinge
        End If
    End If
    'Insert your code for output 4(Stem)
    If (lPartDataBasis = 437) Then
    
    ElseIf (lPartDataBasis = 435) Then
        Dim oWire As IngrGeom3D.ComplexString3d
        Set oWire = New IngrGeom3D.ComplexString3d
        Set oStPoint = New AutoMath.DPosition
        Set oGeomFactory = New IngrGeom3D.GeometryFactory

        'Code for forming Semi-Circular plane
        Set lines = New Collection
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.2 * ClampWidth, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk, ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk, ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, ClampOpening / 2 + 0.2 * ClampSideWallTk)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, ClampOpening / 2 + 0.2 * ClampSideWallTk, -ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, ClampOpening / 2 + 0.2 * ClampSideWallTk)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, ClampOpening / 2 + 0.2 * ClampSideWallTk, -ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -ClampWidth / 2, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk, -0.2 * ClampWidth, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk)
        lines.Add oLine
        
        oStPoint.Set 0.2 * ClampWidth, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk
        Set oWire = PlaceTrCString(oStPoint, lines)
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set oLine = Nothing
        Set oArc = Nothing
        Set lines = Nothing
        Dim oCrossSection As IngrGeom3D.Circle3d
        Set oCrossSection = New Circle3d
        Set oCrossSection = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0.2 * ClampWidth, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.2 * ClampSideWallTk, 1, 0, 0, 0.2 * ClampSideWallTk)
        
        Dim stnorm() As Double
        Dim ednorm() As Double

        Dim Surfset   As IngrGeom3D.IJElements
        Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oWire, oCrossSection, _
                            CircularCorner, 0, stnorm, ednorm, True)
        Dim oSweepCurve As Object
        'Set the Output
        For Each oSweepCurve In Surfset
            m_OutputColl.AddOutput "Sweep_", oSweepCurve
        Next oSweepCurve

        Set oCrossSection = Nothing
        Set Surfset = Nothing
        Set oSweepCurve = Nothing
    Else
        Dim StemStPt As AutoMath.DPosition
        Dim StemEndPt As AutoMath.DPosition
        Set StemStPt = New AutoMath.DPosition
        Set StemEndPt = New AutoMath.DPosition
    
        StemStPt.Set 0, _
                     -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5), _
                     -Sin(HalfOpeningAngle) * InSideRadius - ClampTotalTk
        StemEndPt.Set StemStPt.x, StemStPt.y, StemStPt.z + StemTotalHt
        If (lPartDataBasis = 440) Then
            StemStPt.Set 0, _
                     -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1), _
                     -Sin(HalfOpeningAngle) * InSideRadius - 0.25 * (parClampLength - parClampDiameter * 1.1)
            StemEndPt.Set StemStPt.x, StemStPt.y, StemStPt.z + StemTotalHt
            StemDia = 1.5 * StemDia
        End If

        If (lPartDataBasis = 439 Or lPartDataBasis = 438) Then
            StemStPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 9)
            If (lPartDataBasis = 438) Then
                StemEndPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, OutSideRadius * Sin(PI / 9) + 4.2 * flangeThick
            Else
                StemEndPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, OutSideRadius * Sin(PI / 9) + StemTotalHt
            End If
            StemDia = 0.75 * StemDia
        ElseIf (lPartDataBasis = 431 Or lPartDataBasis = 436 Or lPartDataBasis = 432 Or lPartDataBasis = 433) Then
            StemDia = 1.5 * StemDia
        ElseIf (lPartDataBasis = 434) Then
            StemStPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, OutSideRadius * Sin(PI / 4)
            StemEndPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 9) - StemTotalHt
            StemDia = 1.75 * StemDia
        End If
        If (lPartDataBasis = 435) Then
            StemStPt.Set ClampWidth / 2, -OutSideRadius * Cos(HalfOpeningAngle), -1.01 * ClampOpening / 2
            StemEndPt.Set ClampWidth / 2, -OutSideRadius * Cos(HalfOpeningAngle), 4 * ClampOpening
            Set ObjStem = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, 0.5 * ClampSideWallTk, True)
        Else
            Set ObjStem = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, StemDia, True)
        End If
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
        Set ObjStem = Nothing
        Set StemStPt = Nothing
        Set StemEndPt = Nothing
    End If

    'Insert your code for Output (Nut)
    If (lPartDataBasis = 439 Or lPartDataBasis = 438) Then
        Dim oNut As Object
        Dim oAxisVec As AutoMath.DVector
        Dim oCenter As AutoMath.DPosition
        Set oAxisVec = New AutoMath.DVector
        Set oCenter = New AutoMath.DPosition
        Dim dLineStrPoints(0 To 20) As Double
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(2) = OutSideRadius * Sin(PI / 9)
        dLineStrPoints(3) = -0.7 * flangeThick
        dLineStrPoints(4) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(5) = OutSideRadius * Sin(PI / 9)
        dLineStrPoints(6) = -0.7 * flangeThick
        dLineStrPoints(7) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(8) = OutSideRadius * Sin(PI / 9) + 1.5 * flangeThick
        dLineStrPoints(9) = -flangeThick
        dLineStrPoints(10) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(11) = OutSideRadius * Sin(PI / 9) + 1.5 * flangeThick
        dLineStrPoints(12) = -0.5 * flangeThick
        dLineStrPoints(13) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(14) = OutSideRadius * Sin(PI / 9) + 2.5 * flangeThick
        dLineStrPoints(15) = -0.5 * flangeThick
        dLineStrPoints(16) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(17) = OutSideRadius * Sin(PI / 9) + 3 * flangeThick
        dLineStrPoints(18) = 0
        dLineStrPoints(19) = -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick
        dLineStrPoints(20) = OutSideRadius * Sin(PI / 9) + 3 * flangeThick

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVec.Set 0, 0, 1
        oCenter.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, OutSideRadius * Sin(PI / 9) + 3 * flangeThick
        Set oNut = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenter, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNut
        Set oNut = Nothing
        Set oLineString = Nothing
    End If

    'Insert your code for Output(Wing1)
    If (lPartDataBasis = 439 Or lPartDataBasis = 438) Then
        Dim oWing As IngrGeom3D.ComplexString3d
        Set oWing = New IngrGeom3D.ComplexString3d
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition

        'Code for forming square plane
        Set lines = New Collection
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.3 * flangeThick, -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 2.5 * flangeThick, 0.3 * flangeThick, -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 3 * flangeThick)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.3 * flangeThick, -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 3 * flangeThick, 0.3 * flangeThick + 1.5 * flangeThick * Sin(PI / 6), -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 3 * flangeThick + 1.5 * flangeThick * Cos(PI / 6))
        lines.Add oLine
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, 0.3 * flangeThick + 1.5 * flangeThick * Sin(PI / 6), -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 3 * flangeThick + 1.5 * flangeThick * Cos(PI / 6), 0.3 * flangeThick + 0.25 * flangeThick + 1.5 * flangeThick * Sin(PI / 6), -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 3 * flangeThick + 1.5 * flangeThick * Cos(PI / 6) - 0.25 * flangeThick, 0.3 * flangeThick + 1.5 * flangeThick * Sin(PI / 4), -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 2.5 * flangeThick + 1.5 * flangeThick * Cos(PI / 4))
        lines.Add oArc
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.3 * flangeThick + 1.5 * flangeThick * Sin(PI / 4), -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 2.5 * flangeThick + 1.5 * flangeThick * Cos(PI / 4), 0.3 * flangeThick, -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 2.5 * flangeThick)
        lines.Add oLine

        oStPoint.Set 0.3 * flangeThick, -OutSideRadius * Cos(PI / 9) - 0.5 * flangeThick, OutSideRadius * Sin(PI / 9) + 2.5 * flangeThick
        Set oWing = PlaceTrCString(oStPoint, lines)
        Set oGeomFactory = Nothing
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set oLine = Nothing
        Set oArc = Nothing
        Set lines = Nothing
        'Code for output(Projection)
        Set oAxis = New AutoMath.DVector
        oAxis.Set 0, -1, 0
        Set oProjection = PlaceProjection(m_OutputColl, oWing, oAxis, 0.5 * flangeThick, True)
        m_OutputColl.AddOutput "Wings_", oProjection
        Set oProjection = Nothing
        Set oProjection = PlaceProjection(m_OutputColl, oWing, oAxis, 0.5 * flangeThick, True)
        Set oTransMatrix = New AutoMath.DT4x4
        oTransMatrix.LoadIdentity
        oTransMatrix.IndexValue(0) = -1
        oProjection.Transform oTransMatrix
        m_OutputColl.AddOutput "Wings_", oProjection
        Set oProjection = Nothing
        Set oAxis = Nothing
        Set oWing = Nothing
    End If

    'Insert your code for output 5(Handle)
    If (lPartDataBasis = 430 Or lPartDataBasis = 434 Or lPartDataBasis = 441 Or lPartDataBasis <= 1) Then
        Dim HandleStPt As AutoMath.DPosition
        Dim HandleEndPt As AutoMath.DPosition
        Set HandleStPt = New AutoMath.DPosition
        Set HandleEndPt = New AutoMath.DPosition

        HandleStPt.Set -ScrewHandleLen / 2, _
                     -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5), _
                     -Sin(HalfOpeningAngle) * InSideRadius - ClampTotalTk + StemTotalHt
        HandleEndPt.Set ScrewHandleLen / 2, HandleStPt.y, HandleStPt.z
        If (lPartDataBasis = 434) Then
            HandleStPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 9) - 0.8 * StemTotalHt
            HandleEndPt.Set 0, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick, -OutSideRadius * Sin(PI / 9) - 0.5 * StemTotalHt
            HandleDia = 2.5 * HandleDia
        End If

        Set ObjHandle = PlaceCylinder(m_OutputColl, HandleStPt, HandleEndPt, HandleDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandle
        Set ObjHandle = Nothing
        Set HandleStPt = Nothing
        Set HandleEndPt = Nothing
    End If

    'Insert your code for Output (WingNut)
    If (lPartDataBasis = 431 Or lPartDataBasis = 436 Or lPartDataBasis = 432 Or lPartDataBasis = 433 Or lPartDataBasis = 440) Then
        Dim oWingNut As IngrGeom3D.ComplexString3d
        Set oWingNut = New IngrGeom3D.ComplexString3d
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oStPoint = New AutoMath.DPosition
        If (lPartDataBasis = 440) Then
            'Code for forming Wingnut plane
            Set lines = New Collection
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick)
            lines.Add oLine
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 2 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 4.5 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick)
            lines.Add oArc
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick)
            lines.Add oLine
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 2 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 4.5 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick)
            lines.Add oArc
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick, -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick)
            lines.Add oLine
    
            oStPoint.Set -flangeThick, -OutSideRadius - 0.35 * (parClampLength - parClampDiameter * 1.1) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick
            Set oWingNut = PlaceTrCString(oStPoint, lines)
            For iCount = 1 To lines.Count
                lines.Remove 1
            Next iCount
            Set oLine = Nothing
            Set oArc = Nothing
            Set lines = Nothing
            'Code for output (Projection)
            Set oAxis = New AutoMath.DVector
            oAxis.Set 1, 0, 0
    
            Set oProjection = PlaceProjection(m_OutputColl, oWingNut, oAxis, 2 * flangeThick, True)

        Else
            'Code for forming Wingnut plane
            Set lines = New Collection
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick)
            lines.Add oLine
            If (lPartDataBasis = 436 Or lPartDataBasis = 432 Or lPartDataBasis = 440) Then
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 2 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 4.5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick)
                lines.Add oArc
            Else
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick)
                lines.Add oLine
            End If
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick)
            lines.Add oLine
            If (lPartDataBasis = 436 Or lPartDataBasis = 432 Or lPartDataBasis = 440) Then
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 2 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 4.5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick)
                lines.Add oArc
            Else
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 3 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 5 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick)
                lines.Add oLine
            End If
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening + 3.6 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) - 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick, -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick)
            lines.Add oLine
    
            oStPoint.Set -flangeThick, -OutSideRadius * Cos(HalfOpeningAngle) - ((ScrewSupLen - ClampTotalTk) * 0.5) + 1.5 * flangeThick, -Sin(HalfOpeningAngle) * InSideRadius + ClampTotalTk + ClampOpening - 0.1 * flangeThick
            Set oWingNut = PlaceTrCString(oStPoint, lines)
            For iCount = 1 To lines.Count
                lines.Remove 1
            Next iCount
            Set oLine = Nothing
            Set oArc = Nothing
            Set lines = Nothing
            'Code for output (Projection)
            Set oAxis = New AutoMath.DVector
            oAxis.Set 1, 0, 0
    
            Set oProjection = PlaceProjection(m_OutputColl, oWingNut, oAxis, 2 * flangeThick, True)
        End If
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
        Set oProjection = Nothing
        Set oAxis = Nothing
        Set oGeomFactory = Nothing
        Set oWingNut = Nothing
    End If

    'Insert your code for output (Bolts)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    If (lPartDataBasis = 437) Then
        oStPoint.Set 0, 0.45 * parClampDiameter, 0.65 * parClampDiameter
        oEnPoint.Set 0, -0.6 * parClampDiameter, 0.65 * parClampDiameter
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * ClampWidth, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBolt
        Set oBolt = Nothing
        oStPoint.Set 0, 0.6 * parClampDiameter, -0.65 * parClampDiameter
        oEnPoint.Set 0, -0.45 * parClampDiameter, -0.65 * parClampDiameter
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * ClampWidth, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBolt
    End If
    If (lPartDataBasis = 433) Then
        oStPoint.Set 0, OutSideRadius * Cos(HalfOpeningAngle) + 0.25 * flangeThick, OutSideRadius * Sin(HalfOpeningAngle) + 0.25 * flangeThick
        oEnPoint.Set 0, OutSideRadius * Cos(HalfOpeningAngle) + 0.25 * flangeThick, -OutSideRadius * Sin(HalfOpeningAngle) - 0.25 * flangeThick
        Set oBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.25 * flangeThick, True)
        m_OutputColl.AddOutput "Bolt", oBolt
    End If
    Set oBolt = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Insert your code for Output (Connector)
    If (lPartDataBasis = 435) Then
        Dim oConnector As IngrGeom3D.ComplexString3d
        Set oConnector = New IngrGeom3D.ComplexString3d
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oStPoint = New AutoMath.DPosition

        'Code for forming Connector Plane
        Set lines = New Collection
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle) + 0.3 * ClampOpening, -ClampOpening / 2, -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle) + 0.3 * ClampOpening, ClampOpening / 2)
        lines.Add oLine
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle) + 0.3 * ClampOpening, ClampOpening / 2, -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle) - 0.4 * ClampOpening, 0, -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle) + 0.3 * ClampOpening, -ClampOpening / 2)
        lines.Add oArc

        oStPoint.Set -ClampWidth / 2, OutSideRadius * Cos(HalfOpeningAngle) + 0.3 * ClampOpening, -ClampOpening / 2
        Set oConnector = PlaceTrCString(oStPoint, lines)
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set oLine = Nothing
        Set oArc = Nothing
        Set lines = Nothing
        'Code for output3(Projection)
        Set oAxis = New AutoMath.DVector
        oAxis.Set 1, 0, 0
        Set oProjection = PlaceProjection(m_OutputColl, oConnector, oAxis, ClampWidth, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
        Set oProjection = Nothing
        Set oAxis = Nothing
        Set oGeomFactory = Nothing
        Set oConnector = Nothing
    End If

    'Insert your code for Output (Snap)
    If (lPartDataBasis = 435) Then
        Dim objCone As Object
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        oStPoint.Set 0, -InSideRadius * Cos(HalfOpeningAngle) - ClampOpening / 2, -ClampOpening / 2 - 0.15 * ClampSideWallTk
        oEnPoint.Set -0.3 * ClampWidth, -OutSideRadius * Cos(HalfOpeningAngle) + flangeThick, -ClampOpening / 2 - 0.15 * ClampSideWallTk - 6 * flangeThick
        Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, 0.25 * ClampWidth, 0.1 * flangeThick, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set objCone = Nothing
    End If
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


