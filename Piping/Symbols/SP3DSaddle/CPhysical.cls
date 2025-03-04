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
'   Copyright (c) 2005-08, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CPhysical.cls
'   ProgID:         SP3DSaddle.CSaddle
'   Author:         kkk
'   Creation Date:  Wednesday, November 16 2005
'   Description:
'   For the default case the details were taken from PDS Piping Component Data - Reference Guide at Page no D-75
'   SN=F170 and MC=SAD. Symbol consist of Physical aspect only. Physical aspect is
'   made up of 1.Cylinder as Saddle, 2.Another Cylinder as Saddle Outlet.

'   Physical aspect for the following six part data basis values
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
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  03.DEC.2007     RRK  CR-129358  Enhanced saddle symbol to support clamp-on and snap-on saddles
'  04.APR.2008     RRK  CR-136361  Enhanced saddle symbol for 45 deg oblique saddle(part data basis:926) & 90 deg square saddle(part data basis:925)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick      As Double
    Dim sptOffset        As Double
    Dim flangeDiam       As Double
    Dim flangeDiam2      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    Dim depth3           As Double
    Dim flangeThick2     As Double
    Dim flangeThick3     As Double
    Dim sptOffset2       As Double
    Dim sptOffset3       As Double
    
    
    Dim iOutput     As Integer
    
    Dim objSaddleCylinder As Object
    Dim ObjSaddleOutlet  As Object
    
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
        
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam2, flangeThick3, flangeDiam, sptOffset3, depth3
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
        
    'Insert your code for Saddle geometry depending on the part data basis values and  set the start point, end point and diameter of the outlet cylinder
    Select Case lPartDataBasis
    Case Is <= 1, 505
        Dim PipeCentoOrigin As Double
        PipeCentoOrigin = (parPipeCentertoSaddleEnd - parFacetoCenter) 'L - J
        oStPoint.Set 0, PipeCentoOrigin, 0
        oEnPoint.Set 0, pipeDiam1 / 2 + parSaddleThickness, 0
        Set objSaddleCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dSaddleRadius * 2, True)

    ' Set the output
        m_OutputColl.AddOutput "SaddleCylinder", objSaddleCylinder
        Set objSaddleCylinder = Nothing
        
    'Setting the start and end points of saddle outlet
        oStPoint.Set 0, pipeDiam1 / 2 + parSaddleThickness, 0
        oEnPoint.Set 0, parPipeCentertoSaddleEnd, 0
        'Outlet Diameter is Pipe-2 Diameter + 2* Saddle thickness as given in the PDS Manual
        dCylDiam = pipeDiam2 + 2 * parSaddleThickness
        
    Case 925
        Dim dOuterRadius As Double 'saddle outer radius
        
        'Setting the diameter of the outlet
        dCylDiam = pipeDiam2
        
        'Retrieving the cylinder start position, saddle thickness and saddle radius
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, dOuterRadius)
                    
        'Calling the sub routine to create saddle portion which gets added as  one output to the variable collection
        Call Physical_PDB498_499_504_925_926(parLength, pipeDiam1, dSaddleAngle, dThickness, lPartDataBasis, _
                        m_OutputColl, dFacetoHeaderCenter, pipeDiam2, dOuterRadius)
                        
        'Setting the start and end points of saddle outlet
        oStPoint.Set dCylPos, dCylPos, 0
        oEnPoint.Set dFacetoHeaderCenter / Tan(PI / 4), dFacetoHeaderCenter, 0
    Case 926
        
        'Setting the diameter of the outlet
        dCylDiam = pipeDiam2
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis)

        'Calling the sub routine to create saddle portion which gets added as  one output to the variable collection
        Call Physical_PDB498_499_504_925_926(parLength, pipeDiam1, dSaddleAngle, dThickness, lPartDataBasis, m_OutputColl)
                            
        'Setting the start and end points of saddle outlet
        oStPoint.Set 0, dCylPos, 0
        oEnPoint.Set 0, dFacetoHeaderCenter, 0
        
    Case 498, 499, 504
        
        'Setting the diameter of the outlet
        dCylDiam = GetCylDiam(2, oPartFclt)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis)
                
        'Calling the sub routine to create saddle portion which gets added as  one output to the variable collection
        Call Physical_PDB498_499_504_925_926(parLength, pipeDiam1, dSaddleAngle, dThickness, lPartDataBasis, m_OutputColl)
                                        
        'Setting the start and end points of saddle outlet
        oStPoint.Set 0, dCylPos, 0
        oEnPoint.Set 0, dFacetoHeaderCenter, 0
    Case 502, 503
        
        'Setting the diameter of the outlet
        dCylDiam = GetCylDiam(2, oPartFclt)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, dSaddleRadius)
        
        'Calling a sub routine to create clamp on saddle portion
        'This sub routine adds seven outputs to the variable collection
        Call Physical_PDB502_503(parLength, dSaddleRadius, parWidth, pipeDiam1, m_OutputColl)
        
        'Setting the start and end points of saddle outlet
        oStPoint.Set 0, dCylPos, 0
        oEnPoint.Set 0, dFacetoHeaderCenter, 0
    End Select
    
    'Creating the output for outlet 1 for the saddles using the start and end positions set in the above select-case block
    Set ObjSaddleOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDiam, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddleOutlet
    Set ObjSaddleOutlet = Nothing
    
' Insert code for variable output of second outlet for the partdatabasis value of 503
    If lPartDataBasis = 503 Then 'Clamp-on saddle, double outlet, specified by saddle length, saddle width, saddle radius,
'                                 and branch face to header centerline
        Dim ObjSaddleOutlet2 As Object
        
        'Setting the diameter of the outlet
        dCylDiam = GetCylDiam(3, oPartFclt)
        
        'Retrieving the cylinder start position and saddle thickness
        Call GetCylPos_CylThick_SadlRadius(dCylPos, dThickness, dCylDiam, pipeDiam1, _
                    dSaddleAngle, dFacetoHeaderCenter, depth2, lPartDataBasis, dSaddleRadius)
        
        oStPoint.Set 0, -dCylPos, 0
        oEnPoint.Set 0, -dFacetoHeaderCenter, 0
        
        Set ObjSaddleOutlet2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDiam, True)
        
        ' Set the output
        m_OutputColl.AddOutput "SaddleOutlet2", ObjSaddleOutlet2
        Set ObjSaddleOutlet2 = Nothing
    End If
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
 
 ' Place Nozzle 2
    If lPartDataBasis = 925 Then
        oPlacePoint.Set (dFacetoHeaderCenter) / Tan(PI / 4) + (sptOffset2 - depth2) * Cos(PI / 4), dFacetoHeaderCenter + (sptOffset2 - depth2) * Sin(PI / 4), 0
        oDir.Set Cos(PI / 4), Sin(PI / 4), 0
    Else
        oPlacePoint.Set 0, dFacetoHeaderCenter + sptOffset2 - depth2, 0
        oDir.Set 0, 1, 0
    End If
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

' Creating Nozzle 3 (as a variable output) for the second outlet for the partdatabasis value of 503
    If lPartDataBasis = 503 Then
        oPlacePoint.Set 0, -dFacetoHeaderCenter - sptOffset3 + depth3, 0
        oDir.Set 0, -1, 0
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    ' Set the output
        m_OutputColl.AddOutput "Noz3", objNozzle
    End If
    
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
  
End Sub
'This sub routine is called for part data basis values of 498, 499, 504, 925 and 926,
'to create saddle portion having a circular arc shape. This  adds one output to the variable collection
'Compulsory parameters: Saddle length, header pipe diameter, saddle angle, saddle thickness, part data basis and output collection
'Optional paramaters: Header face to center, outlet diameter and Saddle outer radius. These will be used only for part data basis value 925
Private Sub Physical_PDB498_499_504_925_926(parLength As Double, pipeDiam1 As Double, dSaddleAngle As Double, dThickness As Double, _
                                    lPartDataBasis As Long, m_OutputColl As Object, Optional parFacetoHeaderCenter As Double, Optional pipeDiam2 As Double, Optional dOuterRadius As Double)
    
    Dim objCurvesCollection  As Collection
    Set objCurvesCollection = New Collection
    
    Dim objSaddle As Object
    
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
    'of complex curve are vertical and for remaining cases they are horizontal.
    'For PDB 925, the curve start point along X is different from other cases.
    If lPartDataBasis = 499 Then
        CurvePoint1.Set parLength / 2, -(pipeDiam1 / 2) * Cos(dSaddleAngle / 2), _
                    (pipeDiam1 / 2) * Sin(dSaddleAngle / 2)
        CurvePoint2.Set CurvePoint1.x, CurvePoint1.y - dThickness, CurvePoint1.z
        
        oCentPos.Set parLength / 2, 0, 0
    ElseIf lPartDataBasis = 925 Then
        
        
        'Length of contact between outlet and header
        Dim dLength As Double
        dLength = dOuterRadius * Tan(PI / 4)
        
        'Assumption for saddle curve start point
        Dim dSaddleCurveStPoint As Double
        dSaddleCurveStPoint = (parLength / 2) + dLength
        
        CurvePoint1.Set dSaddleCurveStPoint, (pipeDiam1 / 2) * Cos(dSaddleAngle / 2), _
                        (pipeDiam1 / 2) * Sin(dSaddleAngle / 2)
        CurvePoint2.Set CurvePoint1.x, CurvePoint1.y, CurvePoint1.z + dThickness
        
        oCentPos.Set dSaddleCurveStPoint, 0, 0
    Else
        CurvePoint1.Set parLength / 2, (pipeDiam1 / 2) * Cos(dSaddleAngle / 2), _
                    (pipeDiam1 / 2) * Sin(dSaddleAngle / 2)
        CurvePoint2.Set CurvePoint1.x, CurvePoint1.y, CurvePoint1.z + dThickness
        
        oCentPos.Set parLength / 2, 0, 0
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
    Set objSaddle = PlaceProjection(m_OutputColl, objSaddleOutline, oAxisVec, parLength, True)
    
    ' Set the output
    m_OutputColl.AddOutput "Saddle", objSaddle
    
    Set objSaddle = Nothing
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
Private Sub Physical_PDB502_503(parLength As Double, dSaddleRadius As Double, _
                                parWidth As Double, pipeDiam1 As Double, m_OutputColl As Object)
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    'Create output for Clamp on portion as cylinder
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    oStPoint.Set -parLength / 2, 0, 0
    oEnPoint.Set parLength / 2, 0, 0
    
    Dim objSaddleCylinder As Object
    Set objSaddleCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dSaddleRadius * 2, True)
    
    ' Set the output
    m_OutputColl.AddOutput "SaddleCylinder", objSaddleCylinder
    Set objSaddleCylinder = Nothing
    
    'Create output for Bolt 1 as box
    'Assumptions:1)Length(X) of box=0.2 times saddle length
    '            2)Height(Y) of box=1.8 times saddle radius
    '            3)Width(Z) of box=0.5 times (Saddle Width -Header Diameter)
    Dim objBox1 As Object
    oStPoint.Set 0.4 * parLength, -0.9 * dSaddleRadius, parWidth / 2
    oEnPoint.Set 0.2 * parLength, 0.9 * dSaddleRadius, pipeDiam1 / 2
    
    Set objBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box1", objBox1
    Set objBox1 = Nothing
    
    'Create output for Bolt 2
    Dim objBox2 As Object
    oStPoint.Set -0.4 * parLength, -0.9 * dSaddleRadius, 0.5 * parWidth
    oEnPoint.Set -0.2 * parLength, 0.9 * dSaddleRadius, pipeDiam1 / 2
    
    Set objBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box2", objBox2
    Set objBox2 = Nothing
    
    'Create output for Bolt 3
    Dim objBox3 As Object
    oStPoint.Set 0.4 * parLength, -0.9 * dSaddleRadius, -0.5 * parWidth
    oEnPoint.Set 0.2 * parLength, 0.9 * dSaddleRadius, -pipeDiam1 / 2
    
    Set objBox3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box3", objBox3
    Set objBox3 = Nothing
    
    'Create output for Bolt 4
    Dim objBox4 As Object
    oStPoint.Set -0.4 * parLength, -0.9 * dSaddleRadius, -0.5 * parWidth
    oEnPoint.Set -0.2 * parLength, 0.9 * dSaddleRadius, -pipeDiam1 / 2
    
    Set objBox4 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box4", objBox4
    Set objBox4 = Nothing
    
    'Create output for Connection 1
    'Assumptions:1)Length(X) of box=0.4 times saddle length
    '            2)Height(Y) of box=0.2 times saddle radius
    '            3)Width(Z) of box=0.5 times (Saddle Width -Header Diameter)
    Dim objBox5 As Object
    oStPoint.Set -0.2 * parLength, -0.1 * dSaddleRadius, 0.5 * parWidth
    oEnPoint.Set 0.2 * parLength, 0.1 * dSaddleRadius, pipeDiam1 / 2
    
    Set objBox5 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box5", objBox5
    Set objBox5 = Nothing
    
    'Create output for Connection 2
    Dim objBox6 As Object
    oStPoint.Set -0.2 * parLength, -0.1 * dSaddleRadius, -0.5 * parWidth
    oEnPoint.Set 0.2 * parLength, 0.1 * dSaddleRadius, -pipeDiam1 / 2
    
    Set objBox6 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    m_OutputColl.AddOutput "Box6", objBox6
    Set objBox6 = Nothing

End Sub
'This function is called to get the follwing outputs based on the inputs mentioned below
'Common Outputs:Starting position of the outlet and thickness of the saddle
'Optional Outputs:Saddle outer radius which is required for part data basis 925
'Common Inputs: Outlet Diameter, Header Diameter, Saddle Angle, Face to Header Center and Partdatabasis
Private Function GetCylPos_CylThick_SadlRadius(ByRef dCylPos As Double, ByRef dThickness As Double, _
                ByVal dCylDiam As Double, ByVal pipeDiam1 As Double, ByVal dSaddleAngle As Double, _
                ByVal dFacetoHeaderCenter As Double, ByVal depth, ByVal lPartDataBasis, _
                Optional ByRef dOuterRadius As Double)

    'For PDB values 502 and 503 setting outer radius to saddle radius
    If lPartDataBasis = 502 Or lPartDataBasis = 503 Then
         'Setting the outlet start position
        dCylPos = Sqr(dOuterRadius ^ 2 - (dCylDiam / 2) ^ 2)
    ElseIf lPartDataBasis = 925 Then
        'Assumption:Outer radius of the saddle is 0.6 times header diameter
        dOuterRadius = 0.6 * pipeDiam1

        'Checking to see that outer radius of saddle doesn't exceed face to header center minus depth
        If CmpDblGreaterthan(dOuterRadius, dFacetoHeaderCenter - (depth * Cos(PI / 4))) Then
            dOuterRadius = 0.95 * (dFacetoHeaderCenter - (depth * Cos(PI / 4)))
        End If
        
        'Setting the outlet start position
        dCylPos = dOuterRadius - (0.5 * dCylDiam * Cos(PI / 4))
    Else
        'Assumption:Outer radius of the saddle is 0.6 times header diameter
        dOuterRadius = 0.6 * pipeDiam1
        
        'Checking to see that outer radius of saddle is less than outlet radius
        If CmpDblGreaterthan(dCylDiam / 2, dOuterRadius) Then
            dOuterRadius = 0.55 * dCylDiam
        End If
        
        'Checking to see that outer radius of saddle doesn't exceed face to header center minus depth
        If CmpDblGreaterthan(dOuterRadius, dFacetoHeaderCenter - depth) Then
            dOuterRadius = 0.95 * (dFacetoHeaderCenter - depth)
        End If
        
        'Setting the outlet start position
        dCylPos = Sqr(dOuterRadius ^ 2 - (dCylDiam / 2) ^ 2)
    End If
    
    'Calculating thickness based on partdatabasis values
    If lPartDataBasis = 499 Then
        dThickness = Sqr(dOuterRadius ^ 2 - (0.5 * pipeDiam1 * Sin(dSaddleAngle / 2)) ^ 2) - (0.5 * pipeDiam1 * Cos(dSaddleAngle / 2))
    Else
        dThickness = Sqr(dOuterRadius ^ 2 - (0.5 * pipeDiam1 * Cos(dSaddleAngle / 2)) ^ 2) - (0.5 * pipeDiam1 * Sin(dSaddleAngle / 2))
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
'Function which returns the flange diameter in cases of female ends and pipe diameter in case of other ends
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
