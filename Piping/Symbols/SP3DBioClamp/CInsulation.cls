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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Feb 1 2005
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   SP3DBioClamp Symbol is created with Seven number of out puts.
'
'   Change History:
'   dd.mmm.yyyy     who            change description
'   -----------     -----          ------------------
'   01.Jan.2005     svsmylav       TR-70434: Added insulation aspect class file (CInsulation.cls) to the symbol.
'   08.SEP.2006     KKC            DI-95670 Replace names with initials in all revision history sheets and symbols
'   12.Oct.2007     VRG            CR-127155 Enhanced the existing  BioClamp to support all Tri-Clamp options using PartDataBasis values.
'   15.Nov.2007     VRG            CR-127647 Enhanced the symbol code to support Hinged Clamp for joining Kwik Flanges
'   15.Nov.2007     VRG/PK         CR-127156 Enhanced the symbol code to support Swivel Joint Clamp
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private Const MODULE = "Insulation:"    'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim InsulationDia As Double
    Dim iOutput As Integer
    Dim ObjBodyInsulation As Object
    Dim ObjClampSupportInsulation As Object
    Dim parClampLength As Double
    Dim parClampWidth As Double
    Dim parClampDiameter As Double
    Dim parInsulationThickness As Double
    PI = 4 * Atn(1)

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampDiameter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    If UBound(arrayOfInputs) >= 4 Then   'parClampLength and parClampWidth are used for Hinged Clamp
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
    If (lPartDataBasis = 439) Then
        parClampDiameter = parClampDiameter / 1.1
    End If
    iOutput = 0

    'Insert your code for output 6(Clamp Body Insulation)
    Dim ClampTk As Double, ClampWidth As Double, ClampSideWallTk As Double, ClampTotalTk As Double
    Dim ClampOpening As Double, InSideRadius As Double, OutSideRadius As Double, includAngle As Double
    Dim HalfOpeningAngle As Double, DistalongAxis As Double, ScrewSupLen As Double
    Dim StemTotalHt As Double, StemDia As Double, ScrewHandleLen As Double, HandleDia As Double
    
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
        If (lPartDataBasis = 431 Or lPartDataBasis = 434 Or lPartDataBasis = 436 Or lPartDataBasis = 440) Then   'Partdatabasis value 431 is used for Tri-Clover A13MHM clamp
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
        DistalongAxis = Sqr(Abs((InSideRadius * InSideRadius) - ((ClampOpening * 0.5) * (ClampOpening * 0.5))))
        HalfOpeningAngle = Atn((ClampOpening * 0.5) / DistalongAxis)
        includAngle = HalfOpeningAngle * 2
        
        StemTotalHt = flangeThick * 4 + ClampTotalTk * 2 + ClampOpening
    End If
    If CmpDblGreaterthanOrEqualTo(flangeThick, 0.06 * parClampDiameter) Then
        InsulationDia = 2 * flangeThick + 2.1 * OutSideRadius + 2 * parInsulationThickness
    Else
        InsulationDia = 2 * 0.06 * parClampDiameter + 2.1 * OutSideRadius + 2 * parInsulationThickness
    End If
    If (lPartDataBasis = 440) Then
        InsulationDia = 0.6 * (parClampLength - parClampDiameter * 1.1) + 2 * OutSideRadius + 2 * parInsulationThickness
    End If
    Dim oStartInsulationPos As New AutoMath.DPosition
    Dim oEndInsulationPos As New AutoMath.DPosition

    oStartInsulationPos.Set -ClampWidth / 2 - parInsulationThickness, 0, 0
    oEndInsulationPos.Set ClampWidth / 2 + parInsulationThickness, 0, 0
    Set ObjBodyInsulation = PlaceCylinder(m_OutputColl, oStartInsulationPos, oEndInsulationPos, _
                                          InsulationDia, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyInsulation
    Set ObjBodyInsulation = Nothing

    'Insert your code for output 7(Clamp Support Insulation)
    If (lPartDataBasis = 431 Or lPartDataBasis = 432 Or lPartDataBasis = 433 Or lPartDataBasis = 434 Or lPartDataBasis = 436 Or lPartDataBasis = 438 Or lPartDataBasis = 439 Or lPartDataBasis = 440) Then
        oStartInsulationPos.Set -ClampWidth / 2 - parInsulationThickness, -0.3 * parClampDiameter, -OutSideRadius * Sin(PI / 9) - parInsulationThickness
        oEndInsulationPos.Set ClampWidth / 2 + parInsulationThickness, -parClampDiameter / 2 - (flangeThick * 4 + parInsulationThickness + ((1.1 * parClampDiameter) - (0.9 * parClampDiameter)) / 2), -Sin(PI / 9) * OutSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick + parInsulationThickness
        If (lPartDataBasis = 434) Then
            oStartInsulationPos.Set -ClampWidth / 2 - parInsulationThickness, -parClampDiameter / 2, OutSideRadius * Sin(PI / 4) + parInsulationThickness
            oEndInsulationPos.Set ClampWidth / 2 + parInsulationThickness, -OutSideRadius * Cos(PI / 9) - 0.75 * flangeThick - 2.5 * flangeThick - parInsulationThickness, -OutSideRadius * Sin(PI / 9) - StemTotalHt - parInsulationThickness
        End If
        If (lPartDataBasis = 440) Then
            oStartInsulationPos.Set -ClampWidth / 2 - parInsulationThickness, -parClampDiameter / 2, -OutSideRadius * Sin(PI / 9) - parInsulationThickness
            oEndInsulationPos.Set ClampWidth / 2 + parInsulationThickness, -OutSideRadius * Cos(PI / 9) - 0.7 * (parClampLength - parClampDiameter * 1.1) - parInsulationThickness, -Sin(PI / 9) * OutSideRadius + ClampTotalTk + ClampOpening + 7.5 * flangeThick + parInsulationThickness
        End If
        Set ObjClampSupportInsulation = PlaceBox(m_OutputColl, oStartInsulationPos, oEndInsulationPos)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampSupportInsulation
        
    ElseIf (lPartDataBasis = 437) Then
         Dim oBoltInsulation As Object
         oStartInsulationPos.Set -0.4 * ClampWidth - parInsulationThickness, 0.45 * parClampDiameter + parInsulationThickness, 0.5 * parClampDiameter
         oEndInsulationPos.Set 0.4 * ClampWidth + parInsulationThickness, -0.6 * parClampDiameter - parInsulationThickness, 0.8 * parClampDiameter + parInsulationThickness
         Set ObjClampSupportInsulation = PlaceBox(m_OutputColl, oStartInsulationPos, oEndInsulationPos)
         'Set the output
         iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampSupportInsulation
         Set oStartInsulationPos = New AutoMath.DPosition
         Set oEndInsulationPos = New AutoMath.DPosition
         oStartInsulationPos.Set -0.4 * ClampWidth - parInsulationThickness, -0.45 * parClampDiameter - parInsulationThickness, -0.5 * parClampDiameter
         oEndInsulationPos.Set 0.4 * ClampWidth + parInsulationThickness, 0.6 * parClampDiameter + parInsulationThickness, -0.8 * parClampDiameter - parInsulationThickness
         Set oBoltInsulation = PlaceBox(m_OutputColl, oStartInsulationPos, oEndInsulationPos)
         m_OutputColl.AddOutput "BoltInsulation", oBoltInsulation
         Set oBoltInsulation = Nothing

    ElseIf (lPartDataBasis = 435) Then
         Dim oSnapInsulation As Object
         oStartInsulationPos.Set ClampWidth / 2 + parInsulationThickness, -OutSideRadius * Cos(HalfOpeningAngle) + 0.5 * ClampSideWallTk + parInsulationThickness, -0.8 * ClampOpening - parInsulationThickness
         oEndInsulationPos.Set -ClampWidth / 2 - parInsulationThickness, -OutSideRadius * Cos(HalfOpeningAngle) - 0.5 * ClampSideWallTk - parInsulationThickness, 0.8 * ClampOpening + parInsulationThickness
         Set ObjClampSupportInsulation = PlaceBox(m_OutputColl, oStartInsulationPos, oEndInsulationPos)
         'Set the output
         iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampSupportInsulation
         oStartInsulationPos.Set 0, -OutSideRadius * Cos(HalfOpeningAngle), -ClampOpening / 2 - 0.15 * ClampSideWallTk + parInsulationThickness
         oEndInsulationPos.Set -0.3 * ClampWidth, -OutSideRadius * Cos(HalfOpeningAngle) + flangeThick, -ClampOpening / 2 - 0.15 * ClampSideWallTk - 6 * flangeThick - parInsulationThickness
         Set oSnapInsulation = PlaceCylinder(m_OutputColl, oStartInsulationPos, oEndInsulationPos, 0.25 * ClampWidth + parInsulationThickness, True)
         m_OutputColl.AddOutput "SnapInsulation", oSnapInsulation
         Set oSnapInsulation = Nothing
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 430 Or lPartDataBasis = 441) Then
        Dim dInsBoxLength As Double
        Dim dInsBoxWidth As Double
        Dim dInsBoxHeight As Double
    
        ClampOpening = parClampDiameter * 0.01 * 2
    
        dInsBoxLength = ClampWidth + 2 * parInsulationThickness      'Along X axis
        dInsBoxHeight = flangeThick * 4 + parInsulationThickness + ((1.1 * parClampDiameter) - (0.9 * parClampDiameter)) / 2    'Along Y axis
        dInsBoxWidth = 2 * ClampTotalTk + 2 * parInsulationThickness + ClampOpening  'Along Z axis
    
        oStartInsulationPos.Set -dInsBoxLength / 2, -parClampDiameter / 2, -dInsBoxWidth / 2
        oEndInsulationPos.Set dInsBoxLength / 2, -parClampDiameter / 2 - dInsBoxHeight, dInsBoxWidth / 2 + flangeThick * 4.5    'Used 'flangeThick * 4.5' to account for stem and handle
    
        Set ObjClampSupportInsulation = PlaceBox(m_OutputColl, oStartInsulationPos, oEndInsulationPos)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampSupportInsulation
    End If
    
    Set ObjClampSupportInsulation = Nothing
    Set oStartInsulationPos = Nothing
    Set oEndInsulationPos = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
