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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          SSP
'   Creation Date:  Thursday, Nov 14 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   22 1/2 Degree Elbow.
'   The symbol consists of six outputs, three physical and three insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----                ------------------
'  09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'  08.SEP.2006     KKC                DI-95670 Replace names with initials in all revision history sheets and symbols.
'  18.Jul.2007     dkl                TR-124260 Rectified insulation of the ports to enable placement when flange thickness is zero for value of flange dia.
'                                     Used appropriate CmpdblXXX functions where required.
'  07.Nov.2007     MA                 TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                     in which case small cylinders of negligible thickness will be created.
'  19.Dec.2007     dkl                CR 127556: implemented new part data basis-13,35,36,161,162,163,164,176,177,179 and 180.
'                                     The 22.5 deg elbow is made trimmable. Tangent lengths 1 and 2 are insulated as optional outputs.
'  09.Jan.2008     RRK                CR-134567 Updated elbow symbols to reference PDB values (from 690 to 701) that are specific to 1/16 bends
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI           As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick1     As Double
    Dim cptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    Dim cptOffset2      As Double
    Dim depth2          As Double
    Dim flangeThick2    As Double
    Dim flangeDiam2     As Double
    
    Dim iOutput     As Integer

    Dim objInsulatedBody As Object
    Dim objInsulatedPort1 As Object
    Dim objInsulatedPort2 As Object

    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parSeattoCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTangentLength As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parAngle As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parSeattoCenter = arrayOfInputs(4)
'    parSeat1toCenter = arrayOfInputs(5)
'    parSeat2toCenter = arrayOfInputs(6)
'    parFace1toCenter = arrayOfInputs(7)
'    parFace2toCenter = arrayOfInputs(8)
'    parTangentLength = arrayOfInputs(9)
'    parTangentLength1 = arrayOfInputs(10)
'    parTangentLength2 = arrayOfInputs(11)
'    parAngle = arrayOfInputs(12)
    
    iOutput = 0
    
    Dim dBendAngle22o5Deg As Double
    dBendAngle22o5Deg = PI / 8
    parAngle = dBendAngle22o5Deg
    'Elbow is trimmable as per the input Bend Angle (parAngle). The following check ensures
    'backward compatibility with previous versions.
    If UBound(arrayOfInputs) >= 12 Then parAngle = arrayOfInputs(12)
    
    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
  
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick1, flangeDiam1, _
                                                                    cptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, _
                                                                    cptOffset2, depth2
  
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
' This symbol is based on the following Part data basis values that govern its geometry,
' Symmetrical Elbow-
' Face-to-center dimension basis.
' 1/16 bend, symmetrical, specified by face-to-center
' Seat-to-center dimension (Face-to-center dimension excludes the socket depth1 or thread depth1).
' 1/16 bend, symmetrical, specified by seat-to-center
' Elbow, Symmetrical, specified by Seat to Center and Bend Radius.
' 1/16 bend, symmetrical, specified by seat to center and bend radius
' Elbow, Symmetrical, specified by Face to Center and Bend Radius.
' 1/16 bend, symmetrical, specified by face to center and bend radius
' Elbow, Symmetrical, specified by Seat to Center, Tangent length and Bend Radius.
' 1/16 bend, symmetrical, specified seat to center, tangent length and bend radius
' Elbow, Symmetrical, specified by Face to Center, Tangent length and Bend Radius.
' 1/16 bend, symmetrical, specified by face to center, tangent length and bend radius
' Elbow, Symmetrical, specified by Tangent Length and Bend Radius.
' Elbow, Symmetrical, specified by Face to Center and Tangent Length.

' Asymmetrical Elbow-
' Asymmetrical Face-to-Center dimension basis.
' 1/16 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2
' Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth1 or thread depth1).
' 1/16 bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2
' Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius.
' 1/16 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius
' Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius.
' 1/16 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius
' Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius.
' 1/16 bend, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius
' Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent Length 1, Tangent Length 2, and Bend Radius.
' 1/16 bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius
' Elbow, Asymmetrical, specified by Tangent Length 1, Tangent Length 2 and Bend Radius.
' Elbow, Asymmetrical, specified by Face to Center 1, Tangent Length 1 and Bend Radius.
' Elbow, Asymmetrical, specified by Face to Center 1, Tangent Length 1 and Tangent Length 2.

    Select Case lPartDataBasis
        Case Is <= 1, 10, 690 'Face-to-center dimension basis(10)
                              '1/16 bend, symmetrical, specified by face-to-center(690)
            parFacetoCenter = arrayOfInputs(2)
            dTangentLength1 = flangeThick1  'zero tangent lengths
            dTangentLength2 = flangeThick2
            dElbowRadius = (parFacetoCenter - flangeThick1) / Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 13, 181, 691 'Asymmetrical Face-to-Center dimension basis(13)
                          '1/16 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2(691)
            parFace1toCenter = arrayOfInputs(7)
            parFace2toCenter = arrayOfInputs(8)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dBendAngle22o5Deg / 2)
                dTangentLength1 = flangeThick1
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle22o5Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 35, 693 'Seat-to-center dimension (Face-to-center dimension excludes the socket depth1 or thread depth1)(35)
                     '1/16 bend, symmetrical, specified by seat-to-center(693)
            parSeattoCenter = arrayOfInputs(4)
            dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dBendAngle22o5Deg / 2)
                dTangentLength1 = flangeThick1
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle22o5Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 36, 692 'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth1 or thread depth1)(36)
                     '1/16 bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2(692)
            parSeat1toCenter = arrayOfInputs(5)
            parSeat2toCenter = arrayOfInputs(6)
            dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dBendAngle22o5Deg / 2)
                dTangentLength1 = flangeThick1
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle22o5Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 160 'Elbow, Symmetrical, specified by TangentLength and BendRadius.
            parTangentLength = arrayOfInputs(9)
            dElbowRadius = oPipeComponent.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFace1toCenter = dElbowRadius * Tan(dBendAngle22o5Deg / 2) + parTangentLength
            dFace2toCenter = dFace1toCenter
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dFace1toCenter
                        
        Case 161, 698 'Elbow, Symmetrical, specified by Seat to Center, Tangent length and Bend Radius(161)
                      '1/16 bend, symmetrical, specified seat to center, tangent length and bend radius(698)
            parSeattoCenter = arrayOfInputs(4)
            parTangentLength = arrayOfInputs(9)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 162, 699 'Elbow, Symmetrical, specified Face to Center, Tangent length and Bend Radius(162)
                      '1/16 bend, symmetrical, specified by face to center, tangent length and bend radius(699)
            parFacetoCenter = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(9)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 163, 694  'Elbow, Symmetrical, specified by Seat to Center and Bend Radius(163)
                       '1/16 bend, symmetrical, specified by seat to center and bend radius(694)
            parSeattoCenter = arrayOfInputs(4)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 164, 696  'Elbow, Symmetrical, specified by Face to Center and Bend Radius
                       '1/16 bend, symmetrical, specified by face to center and bend radius
            parFacetoCenter = arrayOfInputs(2)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
                 
        Case 165    'Elbow, Symmetrical, specified by FacetoCenter and TangentLength.
            parFacetoCenter = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(9)
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFacetoCenter - parTangentLength) / Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 170   'Elbow, Asymmetrical, specified by TangentLength1, TangentLength2 and BendRadius
            parTangentLength1 = arrayOfInputs(10)
            parTangentLength2 = arrayOfInputs(11)
            dElbowRadius = oPipeComponent.BendRadius
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + parTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + parTangentLength2
            
        Case 175   'Elbow, Asymmetrical, specified by Face1toCenter, TangentLength1 and BendRadius.
            parFace1toCenter = arrayOfInputs(7)
            parTangentLength1 = arrayOfInputs(10)
            dElbowRadius = oPipeComponent.BendRadius
            dTangentLength1 = parTangentLength1
            dTangentLength2 = flangeThick2
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 176, 695 'Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius(176)
                      '1/16 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius(695)
            parSeat1toCenter = arrayOfInputs(5)
            parSeat2toCenter = arrayOfInputs(6)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
                        
        Case 177, 697 'Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius(177)
                      '1/16 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius(697)
            parFace1toCenter = arrayOfInputs(7)
            parFace2toCenter = arrayOfInputs(8)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 179, 700 'Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius(179)
                      '1/16 bend, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius(700)
            parSeat1toCenter = arrayOfInputs(5)
            parSeat2toCenter = arrayOfInputs(6)
            parTangentLength1 = arrayOfInputs(10)
            parTangentLength2 = arrayOfInputs(11)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 180, 701  'Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent 1, Tangent 2, and Bend Radius(180)
                       '1/16 bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius(701)
            parFace1toCenter = arrayOfInputs(7)
            parFace2toCenter = arrayOfInputs(8)
            parTangentLength1 = arrayOfInputs(10)
            parTangentLength2 = arrayOfInputs(11)
            dElbowRadius = oPipeComponent.BendRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case 185   'Elbow, Asymmetrical, specified by Face1toCenter, TangentLength1 and TangentLength2
            parFace1toCenter = arrayOfInputs(7)
            parTangentLength1 = arrayOfInputs(10)
            parTangentLength2 = arrayOfInputs(11)
            dFace1toCenter = parFace1toCenter
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFace1toCenter - parTangentLength1) / Tan(dBendAngle22o5Deg / 2)
            'Face to centers are recomputed when elbow is trimmable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    
    End Select
    
    Set oPipeComponent = Nothing
 ' Insert your code for output 2(Insulated Body)
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim dInsulationRadius As Double
    dInsulationRadius = pipeDiam / 2 + parInsulationThickness
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -dFace1toCenter + dTangentLength1, dElbowRadius, 0
    CV.Set 0, 0, 1
    'Limit Insulation until parInsulationRadius coincides with the Radius of the Elbow
    If CmpDblLessThanOrEqualTo(dInsulationRadius, dElbowRadius) Then
        'Construct a circle that will be used to create the revolution
        Dim objCircleI   As IngrGeom3D.Circle3d
        Set objCircleI = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             -dFace1toCenter + dTangentLength1, 0, 0, _
                            1, 0, 0, _
                          dInsulationRadius)
         ' Revolve the circle for 90 degree about the Z axis passing by the origin
        Set objInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                        CV, CP, parAngle, True)
        
        Set objCircleI = Nothing
        Set oGeomFactory = Nothing
    Else
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
        Set objInsulatedBody = PlaceRevolution(m_OutputColl, objBsplineI, _
                        CV, CP, parAngle, True)
        
        Set objBsplineI = Nothing
    End If

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
    Set objInsulatedBody = Nothing
    Set CV = Nothing
    Set CP = Nothing
    
' Insulation for Tangent 1.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    If CmpDblGreaterthan(dTangentLength1, flangeThick1 + LINEAR_TOLERANCE) Then
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
        oStPoint.Set dFace2toCenter * Cos(parAngle), dFace2toCenter * Sin(parAngle), 0
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
    oEnPoint.Set -dFace1toCenter + flangeThick1 + parInsulationThickness, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam1, pipeDiam) Then _
            parInsulationDiameter = flangeDiam1 + 2 * parInsulationThickness
    
    If CmpDblEqual(flangeThick1, 0) Then
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
