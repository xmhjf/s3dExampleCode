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
'   Copyright (c) 2005-08, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday, Nov 23 2005
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol places a 5.625 Degree Elbow.
'   The symbol is prepared based on PDS symbol MC=E5, SN=F47.
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   08.SEP.2006     KKC        DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.Aug.2007     VRG        TR-124960 Insulation diameter was computed based on End preparation
'   11.DEC.2007     VRK        CR-131508 Create 5.625 degree elbow symbol for 1/64 bend
'   09.Jan.2008     RRK        CR-134567 Updated elbow symbols to reference PDB values (from 715 to 718) that are specific to 1/64 bends
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI           As Double
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim flangeDiam1      As Double
    Dim depth           As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    
    Dim ObjInsulatedNozzle1 As Object
    Dim ObjInsulatedNozzle2 As Object
    
    Dim iOutput     As Integer
    Dim ObjInsulatedBody As Object
    Dim dInsulationRadius As Double
    Dim objInsulatedPort1 As Object
    Dim objInsulatedPort2 As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parElbowRadius As Double
    Dim parSeattoCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parAngle As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parElbowRadius = arrayOfInputs(4)
'    parSeattoCenter = arrayOfInputs(5)
'    parSeat1toCenter = arrayOfInputs(6)
'    parSeat2toCenter = arrayOfInputs(7)
'    parFace1toCenter = arrayOfInputs(8)
'    parFace2toCenter = arrayOfInputs(9)
'    parAngle = arrayOfInputs(10)
    
    iOutput = 0
    Dim dBendAngle5o625Deg As Double
    dBendAngle5o625Deg = PI / 32
    parAngle = dBendAngle5o625Deg
    'Elbow is trimmable as per the input Bend Angle (parAngle). The following check ensures
    'backward compatibility with previous versions.
    If UBound(arrayOfInputs) >= 10 Then parAngle = arrayOfInputs(10)

    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dFacetoCenterofTrimmedE As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, _
                                                                    cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, _
                                                                    cptOffset2, depth2
                                                                    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

' This symbol is based on the following Part data basis values that govern its geometry,
' Symmetrical Elbow-
' * Face-to-center dimension basis
' * Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)
' * Elbow, Symmetrical, specified by Seat to Center and Bend Radius
' * 1/64 bend, symmetrical, specified by seat to center and bend radius
' * Elbow, Symmetrical, specified by Face to Center and Bend Radius
' * 1/64 bend, symmetrical, specified by face to center and bend radius
' Asymmetrical Elbow-
' * Asymmetrical Face-to-Center dimension basis
' * Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)
' * Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius
' * 1/64 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius
' * Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius
' * 1/64 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius

 Select Case lPartDataBasis
        Case Is <= 1, 10 'Face-to-center dimension basis
            parFacetoCenter = arrayOfInputs(2)
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle5o625Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
          
        Case 13 'Asymmetrical Face-to-Center dimension basis
            parFace1toCenter = arrayOfInputs(8)
            parFace2toCenter = arrayOfInputs(9)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle5o625Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle5o625Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2

        Case 35  'Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).
            parSeattoCenter = arrayOfInputs(5)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle5o625Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle5o625Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 36  'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle5o625Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle5o625Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
           
        Case 163, 715  'Elbow, Symmetrical, specified by Seat to Center and Bend Radius(163)
                       '1/64 bend, symmetrical, specified by seat to center and bend radius(715)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            parSeattoCenter = arrayOfInputs(5)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2

        Case 164, 717 'Elbow, Symmetrical, specified by Face to Center and Bend Radius(164)
                      '1/64 bend, symmetrical, specified by face to center and bend radius(717)
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
     
        Case 176, 716 'Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius(176)
                      '1/64 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius(716)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
          
        Case 177, 718 'Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius(177)
                      '1/64 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius(718)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            parFace1toCenter = arrayOfInputs(8)
            parFace2toCenter = arrayOfInputs(9)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle5o625Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    
    End Select
    
    'Checking for Zero Tangent Length
    If CmpDblLessThan(dTangentLength1, flangeThick) Then dTangentLength1 = flangeThick
    If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2
    
    ' Insert your code for output 2(Insulated Body)
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    dInsulationRadius = pipeDiam / 2 + parInsulationThickness
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    Dim objCircle   As IngrGeom3D.Circle3d
    CP.Set -dFace1toCenter + dTangentLength1, dElbowRadius, 0
    CV.Set 0, 0, 1
    
'  Limit Insulation until dInsulationRadius coincides with the Radius of the Elbow
    If CmpDblLessThanOrEqualTo(dInsulationRadius, dElbowRadius) Then
'   Construct a circle that will be used to create the revolution
        Dim objCircleI   As IngrGeom3D.Circle3d
        Set objCircleI = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             CP.x, 0, 0, _
                            1, 0, 0, _
                          dInsulationRadius)
    ' Revolve the circle for 90 degree about the Z axis passing by the origin
        
        
        Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                        CV, CP, parAngle, True)
                        Set objCircleI = Nothing
        Set oGeomFactory = Nothing
    Else
        'Construct a Bspline limited by the Radius of the elbow
        Dim HalfCord As Double
        HalfCord = Sqr(dInsulationRadius * dInsulationRadius - dElbowRadius * dElbowRadius)

        Dim points(1 To 13)   As New AutoMath.DPosition
        Dim objBsplineI  As IngrGeom3D.BSplineCurve3d
        points(1).Set -dFace1toCenter + dTangentLength1, dElbowRadius, 0
        points(2).Set -dFace1toCenter + dTangentLength1, dElbowRadius, HalfCord
        points(3).Set -dFace1toCenter + dTangentLength1, dElbowRadius / 2, dInsulationRadius
        points(4).Set -dFace1toCenter + dTangentLength1, 0, dInsulationRadius
        points(5).Set -dFace1toCenter + dTangentLength1, -dInsulationRadius / 2, dInsulationRadius
        points(6).Set -dFace1toCenter + dTangentLength1, -dInsulationRadius, dInsulationRadius / 2
        points(7).Set -dFace1toCenter + dTangentLength1, -dInsulationRadius, 0
        points(8).Set -dFace1toCenter + dTangentLength1, -dInsulationRadius, -dInsulationRadius / 2
        points(9).Set -dFace1toCenter + dTangentLength1, -dInsulationRadius / 2, -dInsulationRadius
        points(10).Set -dFace1toCenter + dTangentLength1, 0, -dInsulationRadius
        points(11).Set -dFace1toCenter + dTangentLength1, dElbowRadius / 2, -dInsulationRadius
        points(12).Set -dFace1toCenter + dTangentLength1, dElbowRadius, -HalfCord
        points(13).Set -dFace1toCenter + dTangentLength1, dElbowRadius, 0

        Set objBsplineI = PlaceTrBspline(5, points)
        ' Revolve the Bspline for 90 degree about the Z axis passing by the origin
        Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objBsplineI, _
                        CV, CP, parAngle, True)
        Set objBsplineI = Nothing
    End If

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set CP = Nothing
    Set CV = Nothing
    
' Insulation for Tangent 1.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    If CmpDblGreaterthan(dTangentLength1, flangeThick + LINEAR_TOLERANCE) Then
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set -dFace1toCenter + dTangentLength1, 0, 0
        Dim objInsulatedTangent1 As Object
        Set objInsulatedTangent1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        pipeDiam + parInsulationThickness * 2, True)
            
        ' Set the output
        m_OutputColl.AddOutput "InsulatedTangent1", objInsulatedTangent1
        Set objInsulatedTangent1 = Nothing
    End If

' Insulation for Tangent 2.
    If CmpDblGreaterthan(dTangentLength2, flangeThick2 + LINEAR_TOLERANCE) Then
        oStPoint.Set (dFace2toCenter) * Cos(parAngle), (dFace2toCenter) * Sin(parAngle), 0
        oEnPoint.Set (dFace2toCenter - dTangentLength2) * Cos(parAngle), _
                (dFace2toCenter - dTangentLength2) * Sin(parAngle), 0
        Dim objInsulatedTangent2 As Object
        Set objInsulatedTangent2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        pipeDiam + parInsulationThickness * 2, True)
            
        ' Set the output
        m_OutputColl.AddOutput "InsulatedTangent2", objInsulatedTangent2
        Set objInsulatedTangent2 = Nothing
    End If

' Insert your code for output 2(Insulated Port1)
    Dim parInsulationDiameter As Double
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then _
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    
    If CmpDblEqual(flangeThick, 0) Then
        oEnPoint.Set -dFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
    End If

    Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

'  Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort1
    Set objInsulatedPort1 = Nothing

 ' Insert your code for output 3(Insulated Port2)
    oStPoint.Set dFace2toCenter * Cos(parAngle), dFace2toCenter * Sin(parAngle), 0
    oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness) * Cos(parAngle), _
                (dFace2toCenter - flangeThick2 - parInsulationThickness) * Sin(parAngle), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then _
        parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
    
    If CmpDblEqual(flangeThick2, 0) Then
        oEnPoint.Set (dFace2toCenter - NEGLIGIBLE_THICKNESS) * Cos(parAngle), _
                        (dFace2toCenter - NEGLIGIBLE_THICKNESS) * Sin(parAngle), 0
    End If
    
    Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort2
    Set objInsulatedPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


