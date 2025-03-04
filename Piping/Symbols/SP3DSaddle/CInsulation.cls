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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   ProgID:         SP3DSaddle.CSaddle
'   Author:         RRK
'   Creation Date:  Monday, December 3 2007
'   Description:
'   Insulation aspect for the following six part data basis values
'   1)Part Data Basis value 498-Saddle, specified by saddle length, branch face to header centerline, and angle
'   2)Part Data Basis value 499-Snap-on saddle, specified by saddle length, branch face to header centerline, and angle
'   3)Part Data Basis value 502-Clamp-on saddle, single outlet, specified by saddle length, saddle width, saddle radius, and branch face to header centerline
'   4)Part Data Basis value 503-Clamp-on saddle, double outlet, specified by saddle length, saddle width, saddle radius, and branch face to header centerline
'   5)Part Data Basis value 504-Saddle, specified by saddle length, angle, and branch face to header centerline or seat to header centerline
'   6)Part Data Basis value Default or 505-Saddle, specified by saddle half-length, branch face to header centerline, branch face to saddle end, saddle thickness, and hub thickness
'   7)Saddle, 45 degree oblique, specified by saddle length and branch face to header centerline or seat to header centerline
'   8)Saddle, 90 degree square, specified by saddle length and branch face to header centerline or seat to header centerline

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  03.DEC.2007     RRK  CR-129358  Enhanced saddle symbol to support clamp-on and snap-on saddles
'  04.APR.2008     RRK  CR-136361  Enhanced saddle symbol for 45 deg oblique saddle(part data basis:926) & 90 deg square saddle(part data basis:925)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Insulation:"    'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim flangeThick      As Double
    Dim sptOffset        As Double
    Dim flangeDiam       As Double
    Dim flangeDiam2      As Double
    Dim flangeDiam3      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    Dim depth3           As Double
    Dim flangeThick2     As Double
    Dim flangeThick3     As Double
    Dim sptOffset2       As Double
    Dim sptOffset3       As Double
    
    
    Dim iOutput     As Integer
    
    Dim objInsSaddleCylinder As Object
    Dim ObjInsSaddleOutlet  As Object
    
    Dim parFacetoCenter As Double
    Dim parSaddleThickness As Double
    Dim parPipeCentertoSaddleEnd As Double
    Dim parSaddleRadius As Double
    Dim parFacetoHeaderCenter As Double
    Dim parLength As Double
    Dim parWidth As Double
    Dim parSeattoHeaderSurface As Double
    Dim parSweep As Double
    Dim parSeattoHeaderCenter As Double
    Dim parInsulationThickness As Double
    
    Dim dFacetoHeaderCenter As Double
    Dim dSaddleAngle As Double
    Dim dSaddleRadius As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)                      'J
'    parSaddleThickness = arrayOfInputs(3)                   'K
'    parPipeCentertoSaddleEnd = arrayOfInputs(4)         'L
'    parSaddleRadius = arrayOfInputs(5)                          'M
'    parFacetoHeaderCenter = arrayOfInputs(6)
'    parLength = arrayOfInputs(7)
'    parWidth = arrayOfInputs(8)
'    parSeattoHeaderSurface = arrayOfInputs(9)
'    parSweep = arrayOfInputs(10)
'    parSeattoHeaderCenter = arrayOfInputs(11)
    
    If UBound(arrayOfInputs) > 5 Then
        parInsulationThickness = arrayOfInputs(12)
    End If
    
    'Retrieving PartDataBasis value
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    
    Select Case lPartDataBasis
    Case Is <= 1, 505 'Saddle, specified by saddle half-length, branch face to header centerline,
                      'branch face to saddle end, saddle thickness, and hub thickness
        parFacetoCenter = arrayOfInputs(2)                      'J
        parSaddleThickness = arrayOfInputs(3)                   'K
        parPipeCentertoSaddleEnd = arrayOfInputs(4)         'L
        parSaddleRadius = arrayOfInputs(5)                          'M
        dFacetoHeaderCenter = parPipeCentertoSaddleEnd
        dSaddleRadius = parSaddleRadius
    Case 498 'Saddle, specified by saddle length, branch face to header centerline, and angle
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parSweep = arrayOfInputs(10)
        dFacetoHeaderCenter = parFacetoHeaderCenter
        dSaddleAngle = parSweep
    Case 499 'Snap-on saddle, specified by saddle length, branch face to header centerline, and angle
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parSweep = arrayOfInputs(10)
        dFacetoHeaderCenter = parFacetoHeaderCenter
        dSaddleAngle = parSweep
    Case 502 '502:Clamp-on saddle, single outlet, specified by saddle length, saddle width, saddle radius,
             '    and branch face to header centerline
        parSaddleRadius = arrayOfInputs(5)
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parWidth = arrayOfInputs(8)
        dFacetoHeaderCenter = parFacetoHeaderCenter
        dSaddleRadius = parSaddleRadius
    Case 503 '503:Clamp-on saddle, double outlet, specified by saddle length, saddle width, saddle radius,
             '    and branch face to header centerline
        parSaddleRadius = arrayOfInputs(5)
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parWidth = arrayOfInputs(8)
        dFacetoHeaderCenter = parFacetoHeaderCenter
        
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
        'Setting the default value of saddle radius
        If CmpDblEqual(parSaddleRadius, 0) Then
            dSaddleRadius = SmallerDim(0.9 * (parFacetoHeaderCenter - ReturnMax4(flangeThick3, flangeThick2, depth2, depth3)), _
                                         0.48 * parWidth)
        Else
            dSaddleRadius = parSaddleRadius
        End If
        
    Case 504 'Saddle, specified by saddle length, angle, and branch face to header centerline or
             'seat to header centerline
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parSweep = arrayOfInputs(10)
        parSeattoHeaderCenter = arrayOfInputs(11)
        dSaddleAngle = parSweep
        
        If Not CmpDblEqual(parFacetoHeaderCenter, 0) Then
            dFacetoHeaderCenter = parFacetoHeaderCenter
        ElseIf Not CmpDblEqual(parSeattoHeaderCenter, 0) Then
            dFacetoHeaderCenter = parSeattoHeaderCenter + depth2 - sptOffset2
        End If
    
    Case 925 'Saddle, 45 degree oblique, specified by saddle length and branch face to header centerline or seat to header centerline
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parSeattoHeaderCenter = arrayOfInputs(11)
        dSaddleAngle = PI
        
        If Not CmpDblEqual(parFacetoHeaderCenter, 0) Then
            dFacetoHeaderCenter = parFacetoHeaderCenter
        ElseIf Not CmpDblEqual(parSeattoHeaderCenter, 0) Then
            dFacetoHeaderCenter = parSeattoHeaderCenter + (depth2 - sptOffset2) * Sin(PI / 4)
        End If
    
    Case 926  'Saddle, 90 degree square, specified by saddle length and branch face to header centerline or seat to header centerline
        parFacetoHeaderCenter = arrayOfInputs(6)
        parLength = arrayOfInputs(7)
        parSeattoHeaderCenter = arrayOfInputs(11)
        dSaddleAngle = PI
        
        If Not CmpDblEqual(parFacetoHeaderCenter, 0) Then
            dFacetoHeaderCenter = parFacetoHeaderCenter
        ElseIf Not CmpDblEqual(parSeattoHeaderCenter, 0) Then
            dFacetoHeaderCenter = parSeattoHeaderCenter + depth2 - sptOffset2
        End If
        
    Case Else
        GoTo ErrorLabel:      ' Invalid Part data basis.
    End Select
    
    'Setting the default value of saddle angle to 120 degrees
    If CmpDblEqual(dSaddleAngle, 0) Then dSaddleAngle = 2 * PI / 3
    
    iOutput = 0
  
 ' Insert your code for output 1(SaddleOutlet)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dCylPos As Double
    Dim dCylDiam As Double
    Dim dThickness As Double
    Dim dInsCylDiam As Double
    
    'Insert your code for Saddle geometry depending on the part data basis values
    Select Case lPartDataBasis
    Case Is <= 1, 505
        Dim PipeCentoOrigin As Double
        PipeCentoOrigin = (parPipeCentertoSaddleEnd - parFacetoCenter) 'L - J
        oStPoint.Set 0, PipeCentoOrigin, 0
        oEnPoint.Set 0, pipeDiam1 / 2 + parSaddleThickness, 0
        Set objInsSaddleCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * (dSaddleRadius + parInsulationThickness), True)
        
    ' Set the output
        m_OutputColl.AddOutput "SaddleCylinder", objInsSaddleCylinder
        Set objInsSaddleCylinder = Nothing
        
        oStPoint.Set 0, pipeDiam1 / 2 + parSaddleThickness, 0
        oEnPoint.Set 0, parPipeCentertoSaddleEnd, 0
        'Outlet Diameter is Pipe-2 Diameter + 2* Saddle thickness as given in the PDS Manual
        dInsCylDiam = pipeDiam2 + 2 * parSaddleThickness + 2 * parInsulationThickness
        
    Case 498, 499, 504
        
        'Setting the diameter of the outlet
        dCylDiam = GetCylDiam(2, oPartFclt)
        dInsCylDiam = BodyOD(flangeDiam2, pipeDiam2) + (2 * parInsulationThickness)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, parInsulationThickness)
                    
        'Calling the sub routine to create saddle portion which gets added as  one output to the variable collection
        Call Insulation_PDB498_499_504_925_926(parLength, pipeDiam1, dSaddleAngle, dThickness, _
                        lPartDataBasis, parInsulationThickness, m_OutputColl)
             
        'Setting the start and end points of saddle outlet
        oStPoint.Set 0, dCylPos, 0
        oEnPoint.Set 0, dFacetoHeaderCenter, 0
        
    Case 502, 503
        
        'Setting the diameter of the outlet
        dCylDiam = GetCylDiam(2, oPartFclt)
        dInsCylDiam = BodyOD(flangeDiam2, pipeDiam2) + (2 * parInsulationThickness)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, parInsulationThickness, dSaddleRadius)
                  
        'For part data basis values of 502 and 503, calling a sub routine to create clamp on saddle portion
        'This sub routine adds seven outputs to the variable collection
        Call Insulation_PDB502_503(parLength, dSaddleRadius, parWidth, pipeDiam1, parInsulationThickness, m_OutputColl)
        
        'Setting the start and end points of saddle outlet
        oStPoint.Set 0, dCylPos, 0
        oEnPoint.Set 0, dFacetoHeaderCenter, 0
        
    Case 925
        'Setting the diameter of the outlet
        dCylDiam = GetCylDiam(2, oPartFclt)
        Dim dOuterRadius As Double
        
        dInsCylDiam = BodyOD(flangeDiam2, pipeDiam2) + (2 * parInsulationThickness)
        
        'Retrieving the cylinder start position, saddle thickness and saddle radius
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, parInsulationThickness, dOuterRadius)
                    
        'Calling the sub routine to create saddle portion which gets added as  one output to the variable collection
        Call Insulation_PDB498_499_504_925_926(parLength, pipeDiam1, dSaddleAngle, dThickness, lPartDataBasis, _
                    parInsulationThickness, m_OutputColl, dFacetoHeaderCenter, pipeDiam2, dOuterRadius)
                        
        'Setting the start and end points of saddle outlet
        oStPoint.Set dCylPos, dCylPos, 0
        oEnPoint.Set dFacetoHeaderCenter / Tan(PI / 4), dFacetoHeaderCenter, 0
        
    Case 926
        'Setting the diameter of the outlet
        dCylDiam = pipeDiam2
        dInsCylDiam = BodyOD(flangeDiam2, pipeDiam2) + (2 * parInsulationThickness)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, parInsulationThickness)
             
        'Calling the sub routine to create saddle portion which gets added as  one output to the variable collection
        Call Insulation_PDB498_499_504_925_926(parLength, pipeDiam1, dSaddleAngle, dThickness, lPartDataBasis, _
                    parInsulationThickness, m_OutputColl)
        
        'Setting the start and end points of saddle outlet
        oStPoint.Set 0, dCylPos, 0
        oEnPoint.Set 0, dFacetoHeaderCenter, 0
        
    End Select
    
    'Creating the saddle outlet
    Set ObjInsSaddleOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDiam, True)

' Set the output
    m_OutputColl.AddOutput "InsSaddleOutlet", ObjInsSaddleOutlet
    Set ObjInsSaddleOutlet = Nothing
    
    ' Insert code for variable output of second outlet for the partdatabasis value of 503
    If lPartDataBasis = 503 Then 'Clamp-on saddle, double outlet, specified by saddle length, saddle width, saddle radius,
'                                 and branch face to header centerline
        
        'Setting the diameter of the outlet 2
        dCylDiam = GetCylDiam(3, oPartFclt)
        Dim ObjInsSaddleOutlet2 As Object
        dInsCylDiam = BodyOD(flangeDiam3, pipeDiam3) + (2 * parInsulationThickness)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                dSaddleAngle, dFacetoHeaderCenter, depth3, lPartDataBasis, parInsulationThickness, dSaddleRadius)
        
        oStPoint.Set 0, -dCylPos, 0
        oEnPoint.Set 0, -dFacetoHeaderCenter, 0
        Set ObjInsSaddleOutlet2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDiam, True)
        
        ' Set the output
        m_OutputColl.AddOutput "InsSaddleOutlet2", ObjInsSaddleOutlet2
        Set ObjInsSaddleOutlet2 = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
  
End Sub

'This sub routine is called for part data basis values of 498, 499, 504, 925 and 926,
'to create saddle portion having a circular arc shape. This  adds one output to the variable collection
'Compulsory parameters: Saddle length, header pipe diameter, saddle angle, saddle thickness, part data basis,insulation thickness and output collection
'Optional paramaters: Header face to center, outlet diameter and Saddle outer radius. These will be used only for part data basis value 925
Private Sub Insulation_PDB498_499_504_925_926(parLength As Double, pipeDiam1 As Double, dSaddleAngle As Double, dThickness As Double, lPartDataBasis As Long, _
        parInsulationThickness As Double, m_OutputColl As Object, Optional parFacetoHeaderCenter As Double, Optional pipeDiam2 As Double, Optional dOuterRadius As Double)
                                                            
    Dim objCurvesCollection  As Collection
    Set objCurvesCollection = New Collection
    
    Dim objInsSaddle As Object
    
    Dim CurvePoint1 As AutoMath.DPosition
    Dim CurvePoint2 As AutoMath.DPosition
    Dim CurvePoint3 As AutoMath.DPosition
    Dim CurvePoint4 As AutoMath.DPosition
    
    Set CurvePoint1 = New AutoMath.DPosition
    Set CurvePoint2 = New AutoMath.DPosition
    Set CurvePoint3 = New AutoMath.DPosition
    Set CurvePoint4 = New AutoMath.DPosition

    Dim oCentPos As AutoMath.DPosition
    Set oCentPos = New AutoMath.DPosition
    
    'Setting the points of the complex curve consisting of two arcs and two lines
    
    'Setting the curve points 1 and 2 based on part data basis values, so that for PDB 499 the lines
    'of complex curve are vertical and for remaining cases they are horizontal
    If lPartDataBasis = 499 Then
        CurvePoint1.Set parLength / 2 + parInsulationThickness, -(pipeDiam1 / 2) * Cos(dSaddleAngle / 2), _
                    (pipeDiam1 / 2) * Sin(dSaddleAngle / 2)
        CurvePoint2.Set CurvePoint1.x, CurvePoint1.y - dThickness, CurvePoint1.z
        
        oCentPos.Set parLength / 2 + parInsulationThickness, 0, 0
    ElseIf lPartDataBasis = 925 Then
        
        'Length of contact between outlet and header
        Dim dLength As Double
        dLength = dOuterRadius * Tan(PI / 4)
        
        'Assumption for saddle curve start position
        Dim dSaddleCurveStPoint As Double
        dSaddleCurveStPoint = (parLength / 2) + dLength + parInsulationThickness
        
        CurvePoint1.Set dSaddleCurveStPoint, (pipeDiam1 / 2) * Cos(dSaddleAngle / 2), _
                        (pipeDiam1 / 2) * Sin(dSaddleAngle / 2)
        CurvePoint2.Set CurvePoint1.x, CurvePoint1.y, CurvePoint1.z + dThickness
        
        oCentPos.Set dSaddleCurveStPoint, 0, 0
    Else
        CurvePoint1.Set parLength / 2 + parInsulationThickness, (pipeDiam1 / 2) * Cos(dSaddleAngle / 2), _
                    (pipeDiam1 / 2) * Sin(dSaddleAngle / 2)
        CurvePoint2.Set CurvePoint1.x, CurvePoint1.y, CurvePoint1.z + dThickness
        
        oCentPos.Set parLength / 2 + parInsulationThickness, 0, 0
    End If
    
    CurvePoint3.Set CurvePoint1.x, CurvePoint2.y, -CurvePoint2.z
    CurvePoint4.Set CurvePoint1.x, CurvePoint1.y, -CurvePoint1.z
        
    'Line 1
    Dim oline1   As IngrGeom3D.Line3d
    Set oline1 = New IngrGeom3D.Line3d
        
    Set oline1 = PlaceTrLine(CurvePoint1, CurvePoint2)
    objCurvesCollection.Add oline1
    Set oline1 = Nothing
    
    'Arc 1
    Dim oArc1   As IngrGeom3D.Arc3d
    Set oArc1 = New IngrGeom3D.Arc3d
    
    Dim oAxis   As AutoMath.DVector
    Set oAxis = New AutoMath.DVector
    oAxis.Set -1, 0, 0

    Set oArc1 = PlaceTrArcByCenterNorm(CurvePoint2, CurvePoint3, oCentPos, oAxis)
    objCurvesCollection.Add oArc1
    Set oArc1 = Nothing
    
    'Line 2
    Dim oline2   As IngrGeom3D.Line3d
    Set oline2 = New IngrGeom3D.Line3d

    Set oline2 = PlaceTrLine(CurvePoint3, CurvePoint4)
    
    objCurvesCollection.Add oline2
    Set oline2 = Nothing
    
    'Arc 2
    Dim oArc2   As IngrGeom3D.Arc3d
    Set oArc2 = New IngrGeom3D.Arc3d
    oAxis.Set 1, 0, 0

    Set oArc2 = PlaceTrArcByCenterNorm(CurvePoint4, CurvePoint1, oCentPos, oAxis)
    objCurvesCollection.Add oArc2
    Set oArc2 = Nothing
    
    Dim objSaddleOutline  As IngrGeom3D.ComplexString3d
    Set objSaddleOutline = New IngrGeom3D.ComplexString3d
    
    Set objSaddleOutline = PlaceTrCString(CurvePoint1, objCurvesCollection)

    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set -1, 0, 0
    
    'Projecting the complex curve
    Set objInsSaddle = PlaceProjection(m_OutputColl, objSaddleOutline, oAxisVec, parLength + 2 * parInsulationThickness, True)
    
    ' Set the output
    m_OutputColl.AddOutput "Saddle", objInsSaddle
    
    Set objInsSaddle = Nothing
    Set oAxisVec = Nothing
    Set oCentPos = Nothing
    Set CurvePoint1 = Nothing
    Set CurvePoint2 = Nothing
    Set CurvePoint3 = Nothing
    Set CurvePoint4 = Nothing
    Set objSaddleOutline = Nothing
    Set objCurvesCollection = Nothing
    
    Exit Sub

End Sub
'This sub routine is called for part data basis values of 502 and 503, to create clamp on saddle portion
'This sub routine adds seven outputs to the variable collection
Private Sub Insulation_PDB502_503(parLength As Double, dSaddleRadius As Double, _
                                parWidth As Double, pipeDiam1 As Double, parInsulationThickness As Double, _
                                m_OutputColl As Object)

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    'Create output for Clamp on portion as cylinder
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    oStPoint.Set -parLength / 2 - parInsulationThickness, 0, 0
    oEnPoint.Set parLength / 2 + parInsulationThickness, 0, 0
    
    Dim objInsSaddleCylinder As Object
    Set objInsSaddleCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * (dSaddleRadius + parInsulationThickness), True)
    
    ' Set the output
    m_OutputColl.AddOutput "SaddleCylinder", objInsSaddleCylinder
    Set objInsSaddleCylinder = Nothing
    
    'Create output for Bolt 1 as box
    'Assumptions:1)Length(X) of box=0.2 times saddle length
    '            2)Height(Y) of box=1.8 times saddle radius
    '            3)Width(Z) of box=0.5 times (Saddle Width -Header Diameter)
    Dim objInsBox1 As Object
    oStPoint.Set 0.4 * parLength + parInsulationThickness, -0.9 * dSaddleRadius - parInsulationThickness, parWidth / 2 + parInsulationThickness
    oEnPoint.Set -oStPoint.x, -oStPoint.y, pipeDiam1 / 2
    
    Set objInsBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box1", objInsBox1
    Set objInsBox1 = Nothing
    
    'Create output for Bolt 2
    Dim objInsBox2 As Object
    oStPoint.Set 0.4 * parLength + parInsulationThickness, -0.9 * dSaddleRadius - parInsulationThickness, -parWidth / 2 - parInsulationThickness
    oEnPoint.Set -oStPoint.x, -oStPoint.y, -pipeDiam1 / 2
    
    Set objInsBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box2", objInsBox2
    Set objInsBox2 = Nothing
  

End Sub
'This function is called to get the follwing outputs based on the inputs mentioned below
'Common Outputs:Starting position of the outlet and thickness of the saddle
'Optional Outputs:Saddle outer radius which is required for part data basis 925
'Common Inputs: Outlet Diameter, Header Diameter, Saddle Angle, Face to Header Center, Partdatabasis and Insulation Thickness
Private Function GetCylPos_CylThick_SadlRadius(ByRef dCylPos As Double, ByRef dThickness As Double, _
                                ByVal dCylDiam As Double, ByVal pipeDiam1 As Double, ByVal dSaddleAngle As Double, ByVal dFacetoHeaderCenter As Double, _
                                ByVal depth, ByVal lPartDataBasis, ByVal parInsulationThickness As Double, Optional ByRef dOuterRadius As Double)
                     
    'Considering insulation diameter as maximum of pipe outside diameter and flange or hub outside diameter
    
    'Outer radius of the insulation saddle
    Dim dInsOuterRadius As Double

    'For PDB values 502 and 503 setting outer radius to saddle radius
    If lPartDataBasis = 502 Or lPartDataBasis = 503 Then
         dInsOuterRadius = dOuterRadius + parInsulationThickness
         dCylPos = Sqr(dOuterRadius ^ 2 - ((dCylDiam / 2)) ^ 2)
         
    ElseIf lPartDataBasis = 925 Then
        'Assumption:Outer radius of the saddle is 0.6 times header diameter
        dOuterRadius = 0.6 * pipeDiam1
        dInsOuterRadius = dOuterRadius + parInsulationThickness
        
        'Checking to see that outer radius of saddle doesn't exceed face to header center minus depth
        If CmpDblGreaterthan(dOuterRadius, dFacetoHeaderCenter - (depth * Cos(PI / 4))) Then
            dOuterRadius = 0.95 * (dFacetoHeaderCenter - (depth * Cos(PI / 4)))
            dInsOuterRadius = dOuterRadius + parInsulationThickness
        End If

        'Setting the outlet start position
        dCylPos = dOuterRadius - (0.5 * dCylDiam * Cos(PI / 4))
    Else
        'Assumption:Outer radius of the saddle is 0.6 times header diameter
        dOuterRadius = 0.6 * pipeDiam1
        dInsOuterRadius = 0.6 * pipeDiam1 + parInsulationThickness
        'Checking to see that outer radius of saddle is less than outlet radius
        If CmpDblGreaterthan(dCylDiam / 2, dOuterRadius) Then
            dOuterRadius = 0.55 * dCylDiam
            dInsOuterRadius = 0.55 * dCylDiam + parInsulationThickness
        End If
        
        'Checking to see that outer radius of saddle doesn't exceed face to header center minus depth
        If CmpDblGreaterthan(dOuterRadius, dFacetoHeaderCenter - depth) Then
            dOuterRadius = 0.95 * (dFacetoHeaderCenter - depth)
            dInsOuterRadius = 0.95 * (dFacetoHeaderCenter - depth) + parInsulationThickness
        End If
        
        dCylPos = Sqr(dOuterRadius ^ 2 - ((dCylDiam / 2)) ^ 2)
    End If

    'Calculating thickness based on partdatabasis values
    If lPartDataBasis = 499 Then
        dThickness = Sqr(dInsOuterRadius ^ 2 - (0.5 * pipeDiam1 * Sin(dSaddleAngle / 2)) ^ 2) - (0.5 * pipeDiam1 * Cos(dSaddleAngle / 2))
    Else
        dThickness = Sqr(dInsOuterRadius ^ 2 - (0.5 * pipeDiam1 * Cos(dSaddleAngle / 2)) ^ 2) - (0.5 * pipeDiam1 * Sin(dSaddleAngle / 2))
    End If
            
End Function
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

Private Function GetCylDiam(PortIndex As Integer, ByVal oPartFclt As PartFacelets.IJDPart) As Double

    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Set oCollection = oPartFclt.GetNozzles()
    Set oPipePort = oCollection.Item(PortIndex)

    'Setting the outlet diameter as hub outside diameter when the ends are female and
    'for other cases setting it to pipe (or body) diameter
    If oPipePort.TerminationClass = 15 Then 'For female ends when Termination Class is 'Female'
        GetCylDiam = oPipePort.FlangeOrHubOutsideDiameter
    Else
        GetCylDiam = oPipePort.PipingOutsideDiameter
    End If
    
    Set oPipePort = Nothing
    Set oCollection = Nothing
End Function
