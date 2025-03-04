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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          SSP
'   Creation Date:  Thursday, Nov 14 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   11 1/4 Degree Elbow.
'   The symbol consists of six outputs, three physical and three insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------    -----        ------------------
'   09.Jul.2003    SymbolTeam(India) Copyright Information, Header  is added.
'  08.SEP.2006     KKC               DI-95670  Replace names with initials in all revision history sheets and symbols
'  10.Dec.2007     VRK               CR-CP·131507 Create 11.25 degree elbow symbol for the following options
'                                    i.1/32 bend
'                                    ii.flanged 1/32 bend
'  09.Jan.2008     RRK               CR-134567 Updated elbow symbols to reference PDB values (from 704 to 711) that are specific to 1/32 bends
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double

    Dim iOutput     As Integer
    Dim ObjElbow As Object
    
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
    
    Dim dBendAngle11o25Deg As Double
    dBendAngle11o25Deg = PI / 16
    parAngle = dBendAngle11o25Deg
    'Elbow is trimmable as per the input Bend Angle (parAngle). The following check ensures
    'backward compatibility with previous versions.
    If UBound(arrayOfInputs) >= 10 Then parAngle = arrayOfInputs(10)

    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                    cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, _
                                                                    cptOffset2, depth2
  
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

' This symbol is based on the following Part data basis values that govern its geometry,
' Symmetrical Elbow-
' * Face-to-center dimension basis.
' *1/32 bend, symmetrical, specified by face-to-center
' * Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).
' * 1/32 bend, symmetrical, specified by seat-to-center
' * Elbow, Symmetrical, specified by Seat to Center and Bend Radius.
' * 1/32 bend, symmetrical, specified by seat to center and bend radius
' * Elbow, Symmetrical, specified by Face to Center and Bend Radius.
' * 1/32 bend, symmetrical, specified by face to center and bend radius
' Asymmetrical Elbow-
' * Asymmetrical Face-to-Center dimension basis.
' * 1/32 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2
' * Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).
' * 1/32 bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2
' * Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius.
' * 1/32 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius
' * Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius.
' * 1/32 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius

 Select Case lPartDataBasis
        Case Is <= 1, 10, 705 'Face-to-center dimension basis
                              '1/32 bend, symmetrical, specified by face-to-center
            parFacetoCenter = arrayOfInputs(2)
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle11o25Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2

        Case 13, 707 'Asymmetrical Face-to-Center dimension basis(13)
                     '1/32 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2(707)
            parFace1toCenter = arrayOfInputs(8)
            parFace2toCenter = arrayOfInputs(9)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle11o25Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle11o25Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2

        Case 35, 704 'Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(35)
                     '1/32 bend, symmetrical, specified by seat-to-center
            parSeattoCenter = arrayOfInputs(5)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle11o25Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle11o25Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 36, 706 'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(36)
                     '1/32 bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2(706)
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle11o25Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle11o25Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
             
        Case 163, 708  'Elbow, Symmetrical, specified by Seat to Center and Bend Radius.
                       '1/32 bend, symmetrical, specified by seat to center and bend radius
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            parSeattoCenter = arrayOfInputs(5)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
      
        Case 164, 710 'Elbow, Symmetrical, specified by Face to Center and Bend Radius(164)
                      '1/32 bend, symmetrical, specified by face to center and bend radius(710)
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
        Case 176, 709 'Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius(176)
                      '1/32 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius(709)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
           
        Case 177, 711 'Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius(177)
                      '1/32 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius(711)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            parFace1toCenter = arrayOfInputs(8)
            parFace2toCenter = arrayOfInputs(9)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle11o25Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
      
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    
    End Select
    'Checking for Zero Tangent Length
    If CmpDblLessThan(dTangentLength1, flangeThick) Then dTangentLength1 = flangeThick
    If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2

   ' Insert your code for output 1(Body)
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation

    Dim objCircle   As IngrGeom3D.Circle3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    'Construct a circle that will be used to create the revolution
    
    CP.Set -dFace1toCenter + dTangentLength1, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -dFace1toCenter + dTangentLength1, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                     CV, CP, parAngle, True)

 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set oGeomFactory = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing
        
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim dPorttoOrigin As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    dPorttoOrigin = dFace1toCenter + cptOffset - depth
    oPlacePoint.Set -dPorttoOrigin, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dTangentLength1)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    dPorttoOrigin = dFace2toCenter + cptOffset2 - depth2
    oPlacePoint.Set dPorttoOrigin * Cos(parAngle), dPorttoOrigin * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dTangentLength2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
