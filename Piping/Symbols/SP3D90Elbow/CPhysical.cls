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
'   Author:          NN
'   Creation Date:  Tuesday, Nov 7 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   16.Jan.2006     KKK                 CR-91107  Update elbow symbols to respect Part Data Basis property
'   08.SEP.2006     KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   07-Mar-07       svsmylav            TR-116058: Added new optional input within the part data basis If-Else condition (not to fail the prior to V7 instances)
'   17.Dec-07       KKC                 CR-121408  Added the new part data basis value 449 to support zero radius elbow
'   25.Oct.2007     MA                  CR-127549 Implemented Part Data Basis for the values 10, 13, 35, 36, 161, 162, 163, 164, 176, 177, 179 and 180.
'   07.Jan.2008     PK                  CR-127548 Implemented Part Data Basis for the values from 538 to 583.
'   09.Jan.2008     RRK                 CR-134567 Updated elbow symbols to reference PDB values (from 723 to 730) that are specific to reducing one quarter bends
'   28.Jul.2008     KKC                 CR-145699 Enhance 45 degree and 90 degree elbow symbol to support JIS case
'   02.Sep.2008     MP                  TR-149108  Problems on flipping the 90deg elbow with side outlet (changed the nozzle positions to be in sequence ).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

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
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim cptOffset2      As Double
    Dim depth2          As Double
    
    Dim iOutput                     As Double
    Dim parFacetoCenter             As Double
    Dim parInsulationThickness      As Double
    Dim parElbowRadius              As Double
    Dim parAngle                    As Double
    Dim parFace1toCenter            As Double
    Dim parFace2toCenter            As Double
    Dim parSeat1toCenter            As Double
    Dim parSeat2toCenter            As Double
    Dim parTangentLength1           As Double
    Dim parTangentLength2           As Double
    Dim parSeattoCenter             As Double
    Dim parTangentLength            As Double
    Dim parHeelInletSeattoCenter    As Double
    Dim parHeelInletFacetoCenter    As Double
    Dim parSeat1toSeatHeelInlet     As Double
    Dim parFace1toFaceHeelInlet     As Double
    Dim parCentertoCenterHeelInlet  As Double
    Dim parSeatSideInlettoCenter    As Double
    Dim parFaceSideInlettoCenter    As Double
    Dim parSideInletCentertoSeat    As Double
    Dim parSideInletCentertoFace    As Double
    Dim parSeatSideOutlettoCenter   As Double
    Dim parFaceSideOutlettoCenter   As Double
    Dim parBaseThickness            As Double
    Dim parBaseWidth                As Double
    Dim parCentertoBase             As Double
    Dim parRibThickness             As Double
    
    Dim dTangentLength              As Double
    Dim ArcStPointToCenter          As Double
    Dim dElbowRadius                As Double
    Dim dFace1toCenter              As Double
    Dim dFace2toCenter              As Double
    Dim dTangentLength1             As Double
    Dim dTangentLength2             As Double
    Dim ObjElbow                    As Object
    Dim dSeat1toCenter              As Double
    Dim dSeat2toCenter              As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parElbowRadius = arrayOfInputs(4)
'    parFace1toCenter = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)
'    parSeat1toCenter = arrayOfInputs(7)
'    parSeat2toCenter = arrayOfInputs(8)
'    parTangentLength1 = arrayOfInputs(9)
'    parTangentLength2 = arrayOfInputs(10)
'    parSeattoCenter = arrayOfInputs(11)
'    parTangentLength = arrayOfInputs(12)
'    parAngle = arrayOfInputs(13)
'    parHeelInletSeattoCenter=arrayOfInputs(14)
'    parHeelInletFacetoCenter=arrayOfInputs(15)
'    parSeat1toSeatHeelInlet=arrayOfInputs(16)
'    parFace1toFaceHeelInlet = arrayOfInputs(17)
'    parCentertoCenterHeelInlet = arrayOfInputs(18)
'    parSeatSideInlettoCenter = arrayOfInputs(19)
'    parFaceSideInlettoCenter = arrayOfInputs(20)
'    parSideInletCentertoSeat = arrayOfInputs(21)
'    parSideInletCentertoFace = arrayOfInputs(22)
'    parFaceSideOutlettoCenter = arrayOfInputs(23)
'    parSeatSideOutlettoCenter = arrayOfInputs(24)
'    parBaseThickness = arrayOfInputs(25)
'    parBaseWidth = arrayOfInputs(26)
'    parCentertoBase = arrayOfInputs(27)
'    parRibThickness = arrayOfInputs(28)

    iOutput = 0
    Dim dBendAngle90Deg As Double
    dBendAngle90Deg = PI / 2
    parAngle = dBendAngle90Deg
    
    'Elbow is trimmable as per the input Bend Angle (parAngle). The following check ensures
    'backward compatibility with previous versions.
    If UBound(arrayOfInputs) >= 13 Then parAngle = arrayOfInputs(13)
    
    Dim dYposition As Double
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                            flangeDiam, cptOffset, depth
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                                            flangeDiam, cptOffset2, depth2
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    ' This symbol is based on the following Part data basis values that govern its geometry,
    ' Symmetrical Elbow-
    ' Face-to-center dimension basis.
    ' Seat-to-center dimension (Face-to-center dimension excludes the socket depth1 or thread depth1).
    ' Elbow, Symmetrical, specified by Seat to Center and Bend Radius.
    ' Elbow, Symmetrical, specified by Face to Center and Bend Radius.
    ' Elbow, Symmetrical, specified by Seat to Center, Tangent length and Bend Radius.
    ' Elbow, Symmetrical, specified by Face to Center, Tangent length and Bend Radius.
    ' One-quarter bend, symmetrical, specified seat-to-center, tangent length, and bend radius
    ' One-quarter bend, reducing, symmetrical, specified seat to center, tangent length and bend radius
    ' One-quarter bend, symmetrical, specified face-to-center, tangent length, and bend radius
    ' One-quarter bend, reducing, symmetrical, specified face to center, tangent length and bend radius
    ' One-quarter bend, symmetrical, specified seat-to-center and bend radius
    ' One-quarter bend, symmetrical, specified face-to-center and bend radius
    ' One-quarter bend, symmetrical, specified by seat-to-center
    ' One-quarter bend, reducing, symmetrical, specified by seat-to-center
    ' One-quarter bend, symmetrical, specified by face-to-center
    ' One-quarter bend, reducing, symmetrical, specified by face-to-center
    ' One-quarter bend with low heel inlet, symmetrical, specified by seat-to-center and heel inlet seat-to-center
    ' One-quarter bend with low heel inlet, symmetrical, specified by face-to-center and heel inlet face-to-center
    ' One-quarter bend with low heel inlet, symmetrical, specified by seat-to-center and seat 1 to seat heel inlet
    ' One-quarter bend with low heel inlet, symmetrical, specified by face-to-center and face 1 to face heel inlet
    ' One-quarter bend with low heel inlet, symmetrical, specified by seat-to-center, tangent length, bend radius, and seat 1 to seat heel inlet
    ' One-quarter bend with low heel inlet, symmetrical, specified by face-to-center, tangent length, bend radius, and face 1 to face heel inlet
    ' One-quarter bend with high heel inlet, symmetrical, specified by seat-to-center, tangent length, bend radius, seat 1 to seat heel inlet, and center to center heel inlet
    ' One-quarter bend with high heel inlet, symmetrical, specified by face-to-center, tangent length, bend radius, face 1 to face heel inlet, and center to center heel inlet
    ' One-quarter bend with high heel inlet, symmetrical, specified by seat-to-center, seat 1 to seat heel inlet, and center to center heel inlet
    ' One-quarter bend with high heel inlet, symmetrical, specified by face-to-center, face 1 to face heel inlet, and center to center heel inlet
    ' One-quarter bend with side inlet, symmetrical, specified by seat-to-center, seat side inlet to center, and side inlet center to seat
    ' One-quarter bend with side inlet, symmetrical, specified by face-to-center, face side inlet to center, and side inlet center to face
    ' 90 degree elbow with side outlet, symmetrical, specified by face-to-center
    ' 90 degree elbow with side outlet, symmetrical, specified by seat-to-center
    ' Base bend, symmetrical, specified by face-to-center, tangent length, bend radius, center-to-base, base thickness, base width, and rib thickness
    ' Base bend, symmetrical, specified by seat-to-center, tangent length, bend radius, center-to-base, base thickness, base width, and rib thickness
    
    ' Asymmetrical Elbow-
    ' Asymmetrical Face-to-Center dimension basis.
    ' Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth1 or thread depth1).
    ' Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius.
    ' Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius.
    ' Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius.
    ' Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent Length 1, Tangent Length 2, and Bend Radius.
    ' One-quarter bend, asymmetrical, specified seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, and bend radius
    ' One-quarter bend, reducing, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius
    ' One-quarter bend, asymmetrical, specified face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius
    ' One-quarter bend, reducing, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius
    ' One-quarter bend, asymmetrical, specified seat-to-center 1, seat-to-center 2, and bend radius
    ' One-quarter bend, asymmetrical, specified face-to-center 1, face-to-center 2, and bend radius
    ' One-quarter bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2
    ' One-quarter bend, reducing, asymmetrical, specified by seat-to-center 1 and seat-to-center 2
    ' One-quarter bend, asymmetrical, specified by face-to-center 1 and face-to-center 2
    ' One-quarter bend, reducing, asymmetrical, specified by face-to-center 1 and face-to-center 2
    ' One-quarter bend with low heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, and heel inlet seat-to-center
    ' One-quarter bend with low heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, and heel inlet face-to-center
    ' One-quarter bend with low heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, and seat 1 to seat heel inlet
    ' One-quarter bend with low heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, and face 1 to face heel inlet
    ' One-quarter bend with low heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, bend radius, and seat 1 to seat heel inlet
    ' One-quarter bend with low heel inlet, asymmetrical, specifid by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, bend radius, and face 1 to face heel inlet
    ' One-quarter bend with high heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, bend radius, seat 1 to seat heel inlet, and center to center heel inlet
    ' One-quarter bend with high heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, bend radius, face 1 to face heel inlet, and center to center heel inlet
    ' One-quarter bend with high heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, seat 1 to seat heel inlet, and center to center heel inlet
    ' One-quarter bend with high heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, face 1 to face heel inlet, and center to center heel inlet
    ' One-quarter bend with side inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, seat side inlet to center, and side inlet center to seat
    ' One-quarter bend with side inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, face side inlet to center, and side inlet center to face
    ' 90 degree elbow with side outlet, asymmetrical, specified by face-to-center 1, face-to-center 2, where side outlet equals face-to-center 1
    ' 90 degree elbow with side outlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, where side outlet equals seat-to-center 1
    ' 90 degree elbow with side outlet, asymmetrical, specified by face-to-center 1, face-to-center 2, where side outlet equals face-to-center 2
    ' 90 degree elbow with side outlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, where side outlet equals seat-to-center 2
    ' Base bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, bend radius, center-to-base, base thickness, base width, and rib thickness
    ' Base bend, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, bend radius, center-to-base, base thickness, base width, and rib thickness

    If (lPartDataBasis <= 1) Then
            parFacetoCenter = arrayOfInputs(2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            ArcStPointToCenter = parFacetoCenter
            dYposition = parFacetoCenter
    
    ElseIf (lPartDataBasis = 449) Then
            parFacetoCenter = arrayOfInputs(2)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - (pipeDiam / 2)
            dTangentLength2 = dFace2toCenter - (pipeDiam / 2)
            ArcStPointToCenter = pipeDiam / 2
            dYposition = pipeDiam / 2
    
    ElseIf (lPartDataBasis = 10) Or (lPartDataBasis = 547) Or (lPartDataBasis = 726) Then 'Face-to-center dimension basis(10)
                                                                'One-quarter bend, symmetrical, specified by face-to-center(547)
                                                                'One-quarter bend, reducing, symmetrical, specified by face-to-center(726)
            parFacetoCenter = arrayOfInputs(2)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
    ElseIf (lPartDataBasis = 13) Or (lPartDataBasis = 549) Or (lPartDataBasis = 725) Then 'Asymmetrical Face-to-Center dimension basis(13)
                                                                'One-quarter bend, asymmetrical, specified by face-to-center 1 and face-to-center 2(549)
                                                                'One-quarter bend, reducing, asymmetrical, specified by face-to-center 1 and face-to-center 2(725)
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
      
    ElseIf (lPartDataBasis = 35) Or (lPartDataBasis = 546) Or (lPartDataBasis = 724) Then 'Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(35)
                                                                'One-quarter bend, symmetrical, specified by seat-to-center(546)
                                                                'One-quarter bend, reducing, symmetrical, specified by seat-to-center(724)
            parSeattoCenter = arrayOfInputs(11)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 36) Or (lPartDataBasis = 548) Or (lPartDataBasis = 723) Then 'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(36)
                                                               'One-quarter bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2(548)
                                                               'One-quarter bend, reducing, asymmetrical, specified by seat-to-center 1 and seat-to-center 2(723)
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 161) Or (lPartDataBasis = 540) Or (lPartDataBasis = 728) Then 'Elbow, Symmetrical, specified Seat to Center, Tangent length and Bend Radius.(161)
                                                                 'One-quarter bend, symmetrical, specified seat-to-center, tangent length, and bend radius(540)
                                                                 'One-quarter bend, reducing, symmetrical, specified seat to center, tangent length and bend radius(728)
            parElbowRadius = arrayOfInputs(4)
            parSeattoCenter = arrayOfInputs(11)
            parTangentLength = arrayOfInputs(12)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 162) Or (lPartDataBasis = 541) Or (lPartDataBasis = 729) Then 'Elbow, Symmetrical, specified Face to Center, Tangent length and Bend Radius(162)
                                                                 'One-quarter bend, symmetrical, specified face-to-center, tangent length, and bend radius(541)
                                                                 'One-quarter bend, reducing, symmetrical, specified face to center, tangent length and bend radius(729)
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            parTangentLength = arrayOfInputs(12)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 163) Or (lPartDataBasis = 542) Then 'Elbow, Symmetrical, specified by Seat to Center and Bend Radius.(162)
                                                                 'One-quarter bend, symmetrical, specified seat-to-center and bend radius(542)
            parElbowRadius = arrayOfInputs(4)
            parSeattoCenter = arrayOfInputs(11)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 164) Or (lPartDataBasis = 55) Or (lPartDataBasis = 543) Then 'Elbow, Symmetrical, specified by Face to Center and Bend Radius.(164)
                                                                                          'One-quarter bend, symmetrical, specified face-to-center and bend radius(543)
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            ArcStPointToCenter = dFace1toCenter - dTangentLength1
            dYposition = dElbowRadius
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
   
    ElseIf (lPartDataBasis = 176) Or (lPartDataBasis = 544) Then 'Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius.(176)
                                                                 'One-quarter bend, asymmetrical, specified seat-to-center 1, seat-to-center 2, and bend radius(544)
            parElbowRadius = arrayOfInputs(4)
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
                        
    ElseIf (lPartDataBasis = 177) Or (lPartDataBasis = 545) Then 'Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius.(177)
                                                                 'One-quarter bend, asymmetrical, specified face-to-center 1, face-to-center 2, and bend radius(545)
            parElbowRadius = arrayOfInputs(4)
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 179) Or (lPartDataBasis = 538) Or (lPartDataBasis = 727) Then 'Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius(179)
                                                                 'One-quarter bend, asymmetrical, specified seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, and bend radius(538)
                                                                 'One-quarter bend, reducing, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius(727)
            parElbowRadius = arrayOfInputs(4)
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
    
    ElseIf (lPartDataBasis = 180) Or (lPartDataBasis = 539) Or (lPartDataBasis = 730) Then 'Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent length 1, Tangent length 2, and Bend Radius(180)
                                                                 'One-quarter bend, asymmetrical, specified face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius(539)
                                                                 'One-quarter bend, reducing, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius(730)
            parElbowRadius = arrayOfInputs(4)
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
            
    ElseIf (lPartDataBasis = 550) Then  'One-quarter bend with low heel inlet, symmetrical, specified by seat-to-center and heel inlet seat-to-center
            parSeattoCenter = arrayOfInputs(11)
            parHeelInletSeattoCenter = arrayOfInputs(14)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 551) Then  'One-quarter bend with low heel inlet, symmetrical, specified by face-to-center and heel inlet face-to-center
            parFacetoCenter = arrayOfInputs(2)
            parHeelInletFacetoCenter = arrayOfInputs(15)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2

    ElseIf (lPartDataBasis = 552) Then  'One-quarter bend with low heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, and heel inlet seat-to-center
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parHeelInletSeattoCenter = arrayOfInputs(14)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 553) Then  'One-quarter bend with low heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, and heel inlet face-to-center
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parHeelInletFacetoCenter = arrayOfInputs(15)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 554) Then  'One-quarter bend with low heel inlet, symmetrical, specified by seat-to-center and seat 1 to seat heel inlet
            parSeattoCenter = arrayOfInputs(11)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            dSeat1toCenter = parSeattoCenter
            dSeat2toCenter = parSeattoCenter

    ElseIf (lPartDataBasis = 555) Then  'One-quarter bend with low heel inlet, symmetrical, specified by face-to-center and face 1 to face heel inlet
            parFacetoCenter = arrayOfInputs(2)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2

    ElseIf (lPartDataBasis = 556) Then  'One-quarter bend with low heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, and seat 1 to seat heel inlet
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            dSeat1toCenter = parSeat1toCenter
            dSeat2toCenter = parSeat2toCenter
            
    ElseIf (lPartDataBasis = 557) Then  'One-quarter bend with low heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, and face 1 to face heel inlet
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 558) Then  'One-quarter bend with low heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, bend radius, and seat 1 to seat heel inlet
            parElbowRadius = arrayOfInputs(4)
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dSeat1toCenter = parSeat1toCenter
            dSeat2toCenter = parSeat2toCenter
            
    ElseIf (lPartDataBasis = 559) Then  'One-quarter bend with low heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, bend radius, and face 1 to face heel inlet
            parElbowRadius = arrayOfInputs(4)
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)

    ElseIf (lPartDataBasis = 560) Then  'One-quarter bend with low heel inlet, symmetrical, specified by seat-to-center, tangent length, bend radius, and seat 1 to seat heel inlet
            parElbowRadius = arrayOfInputs(4)
            parSeattoCenter = arrayOfInputs(11)
            parTangentLength = arrayOfInputs(12)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dSeat1toCenter = parSeattoCenter
            dSeat2toCenter = parSeattoCenter

    ElseIf (lPartDataBasis = 561) Then  'One-quarter bend with low heel inlet, symmetrical, specified by face-to-center, tangent length, bend radius, and face 1 to face heel inlet
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            parTangentLength = arrayOfInputs(12)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)

    ElseIf (lPartDataBasis = 562) Then  'One-quarter bend with high heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, bend radius, seat 1 to seat heel inlet, and center to center heel inlet
            parElbowRadius = arrayOfInputs(4)
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dSeat1toCenter = parSeat1toCenter
            dSeat2toCenter = parSeat2toCenter

    ElseIf (lPartDataBasis = 563) Then  'One-quarter bend with high heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, bend radius, face 1 to face heel inlet, and center to center heel inlet
            parElbowRadius = arrayOfInputs(4)
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)

    ElseIf (lPartDataBasis = 564) Then  'One-quarter bend with high heel inlet, symmetrical, specified by seat-to-center, tangent length, bend radius, seat 1 to seat heel inlet, and center to center heel inlet
            parElbowRadius = arrayOfInputs(4)
            parSeattoCenter = arrayOfInputs(11)
            parTangentLength = arrayOfInputs(12)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dSeat1toCenter = parSeattoCenter
            dSeat2toCenter = parSeattoCenter

    ElseIf (lPartDataBasis = 565) Then  'One-quarter bend with high heel inlet, symmetrical, specified by face-to-center, tangent length, bend radius, face 1 to face heel inlet, and center to center heel inlet
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            parTangentLength = arrayOfInputs(12)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)

    ElseIf (lPartDataBasis = 566) Then  'One-quarter bend with high heel inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, seat 1 to seat heel inlet, and center to center heel inlet
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            dSeat1toCenter = parSeat1toCenter
            dSeat2toCenter = parSeat2toCenter
            
    ElseIf (lPartDataBasis = 567) Then  'One-quarter bend with high heel inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, face 1 to face heel inlet, and center to center heel inlet
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 568) Then  'One-quarter bend with high heel inlet, symmetrical, specified by seat-to-center, seat 1 to seat heel inlet, and center to center heel inlet
            parSeattoCenter = arrayOfInputs(11)
            parSeat1toSeatHeelInlet = arrayOfInputs(16)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            dSeat1toCenter = parSeattoCenter
            dSeat2toCenter = parSeattoCenter
            
    ElseIf (lPartDataBasis = 569) Then  'One-quarter bend with high heel inlet, symmetrical, specified by face-to-center, face 1 to face heel inlet, and center to center heel inlet
            parFacetoCenter = arrayOfInputs(2)
            parFace1toFaceHeelInlet = arrayOfInputs(17)
            parCentertoCenterHeelInlet = arrayOfInputs(18)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2

    ElseIf (lPartDataBasis = 570) Then  'One-quarter bend with side inlet, symmetrical, specified by seat-to-center, seat side inlet to center, and side inlet center to seat
            parSeattoCenter = arrayOfInputs(11)
            parSeatSideInlettoCenter = arrayOfInputs(19)
            parSideInletCentertoSeat = arrayOfInputs(21)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            dSeat1toCenter = parSeattoCenter
            dSeat2toCenter = parSeattoCenter
        
    ElseIf (lPartDataBasis = 571) Then  'One-quarter bend with side inlet, symmetrical, specified by face-to-center, face side inlet to center, and side inlet center to face
            parFacetoCenter = arrayOfInputs(2)
            parFaceSideInlettoCenter = arrayOfInputs(20)
            parSideInletCentertoFace = arrayOfInputs(22)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2

    ElseIf (lPartDataBasis = 572) Then  'One-quarter bend with side inlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, seat side inlet to center, and side inlet center to seat
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parSeatSideInlettoCenter = arrayOfInputs(19)
            parSideInletCentertoSeat = arrayOfInputs(21)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If
            dSeat1toCenter = parSeat1toCenter
            dSeat2toCenter = parSeat2toCenter
            
    ElseIf (lPartDataBasis = 573) Then  'One-quarter bend with side inlet, asymmetrical, specified by face-to-center 1, face-to-center 2, face side inlet to center, and side inlet center to face
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFaceSideInlettoCenter = arrayOfInputs(20)
            parSideInletCentertoFace = arrayOfInputs(22)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 574) Then  '90 degree elbow with side outlet, symmetrical, specified by face-to-center
            parFacetoCenter = arrayOfInputs(2)
            parFaceSideOutlettoCenter = arrayOfInputs(23)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
            dTangentLength1 = flangeThick
            dTangentLength2 = flangeThick2
            If CmpDblEqual(parFaceSideOutlettoCenter, 0) Then
                parFaceSideOutlettoCenter = parFacetoCenter
            End If
    ElseIf (lPartDataBasis = 575) Then  '90 degree elbow with side outlet, symmetrical, specified by seat-to-center
            parSeattoCenter = arrayOfInputs(11)
            parSeatSideOutlettoCenter = arrayOfInputs(24)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 576) Then  '90 degree elbow with side outlet, asymmetrical, specified by face-to-center 1, face-to-center 2, where side outlet equals face-to-center 1
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFaceSideOutlettoCenter = arrayOfInputs(23)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 577) Then  '90 degree elbow with side outlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, where side outlet equals seat-to-center 1
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parSeatSideOutlettoCenter = arrayOfInputs(24)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 578) Then  '90 degree elbow with side outlet, asymmetrical, specified by face-to-center 1, face-to-center 2, where side outlet equals face-to-center 2
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFaceSideOutlettoCenter = arrayOfInputs(23)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 579) Then  '90 degree elbow with side outlet, asymmetrical, specified by seat-to-center 1, seat-to-center 2, where side outlet equals seat-to-center 2
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parSeatSideOutlettoCenter = arrayOfInputs(24)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle90Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle90Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            End If

    ElseIf (lPartDataBasis = 580) Then  'Base bend, symmetrical, specified by face-to-center, tangent length, bend radius, center-to-base, base thickness, base width, and rib thickness
            parElbowRadius = arrayOfInputs(4)
            parSeattoCenter = arrayOfInputs(11)
            parTangentLength = arrayOfInputs(12)
            parBaseThickness = arrayOfInputs(25)
            parBaseWidth = arrayOfInputs(26)
            parCentertoBase = arrayOfInputs(27)
            parRibThickness = arrayOfInputs(28)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)

    ElseIf (lPartDataBasis = 581) Then  'Base bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, bend radius, center-to-base, base thickness, base width, and rib thickness
            parFacetoCenter = arrayOfInputs(2)
            parElbowRadius = arrayOfInputs(4)
            parTangentLength = arrayOfInputs(12)
            parBaseThickness = arrayOfInputs(25)
            parBaseWidth = arrayOfInputs(26)
            parCentertoBase = arrayOfInputs(27)
            parRibThickness = arrayOfInputs(28)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dElbowRadius = parElbowRadius
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)

    ElseIf (lPartDataBasis = 582) Then  'Base bend, symmetrical, specified by seat-to-center, tangent length, bend radius, center-to-base, base thickness, base width, and rib thickness
            parElbowRadius = arrayOfInputs(4)
            parSeat1toCenter = arrayOfInputs(7)
            parSeat2toCenter = arrayOfInputs(8)
            parTangentLength = arrayOfInputs(12)
            parBaseThickness = arrayOfInputs(25)
            parBaseWidth = arrayOfInputs(26)
            parCentertoBase = arrayOfInputs(27)
            parRibThickness = arrayOfInputs(28)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            
    ElseIf (lPartDataBasis = 583) Then  'Base bend, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent length 1, tangent length 2, bend radius, center-to-base, base thickness, base width, and rib thickness
            parElbowRadius = arrayOfInputs(4)
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parTangentLength = arrayOfInputs(12)
            parBaseThickness = arrayOfInputs(25)
            parBaseWidth = arrayOfInputs(26)
            parCentertoBase = arrayOfInputs(27)
            parRibThickness = arrayOfInputs(28)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle90Deg / 2)
            
    ElseIf (lPartDataBasis = 182) Then ' Elbow, asymmetrical, specified by tangent 1, face-to-center 2
        parFace2toCenter = arrayOfInputs(6)
        parTangentLength1 = arrayOfInputs(9)
        dElbowRadius = parFace2toCenter
        dFace1toCenter = parFace2toCenter + parTangentLength1
        dFace2toCenter = parFace2toCenter
        dTangentLength1 = parTangentLength1
        dTangentLength2 = 0
        ArcStPointToCenter = dFace2toCenter
        dYposition = dElbowRadius
            
    Else
            GoTo ErrorLabel:
    End If
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Select Case lPartDataBasis

        Case Is <= 1, 55, 449, 182
  
            'Insert your code for output 1(Elbow)
            'Construct a circle that will be used to create the revolution
            Dim objCircle   As IngrGeom3D.Circle3d
            Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
                            
            'Revolve the circle about the Z axis passing by the origin
            If CmpDblLessThan(dYposition, pipeDiam / 2) Then dYposition = pipeDiam / 2
            Dim CP As New AutoMath.DPosition 'revolution center point
            Dim CV As New AutoMath.DVector 'rotation vector for rotation
            CP.Set -ArcStPointToCenter, dYposition, 0
            CV.Set 0, 0, 1
            
            Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                        CV, CP, parAngle, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
            Set ObjElbow = Nothing
            Set CP = Nothing
            Set CV = Nothing
            Dim sp3dElem1 As IJDObject
            Set sp3dElem1 = objCircle
            Set objCircle = Nothing
            sp3dElem1.Remove
            
                                    
             'Place Nozzle 1
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            
            oPlacePoint.Set -dFace1toCenter - cptOffset + depth, 0, 0
            oDir.Set -1, 0, 0
        
            If CmpDblLessThan(dTangentLength1, flangeThick) Then dTangentLength1 = flangeThick
        
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                                oPlacePoint, dTangentLength1)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
            
            'Place Nozzle 2
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            
            oPlacePoint.Set (dFace2toCenter + cptOffset2 - depth2) * Cos(parAngle), (dFace2toCenter + cptOffset2 - depth2) * Sin(parAngle), 0
            oDir.Set Cos(parAngle), Sin(parAngle), 0
            If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2
            
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                                oPlacePoint, dTangentLength2)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
            Set geomFactory = Nothing
                
                        
        Case 10, 13, 35, 36, 161, 162, 163, 164, 176, 177, 179, 180, 538 To 583, 723 To 730
        
            Dim objArcPort1_1 As IngrGeom3D.EllipticalArc3d
            Set geomFactory = New IngrGeom3D.GeometryFactory
            Set objArcPort1_1 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                    Nothing, _
                                    -dFace1toCenter + dTangentLength1, 0, 0, _
                                    1, 0, 0, _
                                    0, 0, pipeDiam / 2, _
                                    1, 0, PI)
            Dim oCurvesColl As Collection
            Set oCurvesColl = New Collection
            oCurvesColl.Add objArcPort1_1
            Set objArcPort1_1 = Nothing
          
            Dim objArcPort1_2 As IngrGeom3D.EllipticalArc3d
            Set objArcPort1_2 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                    Nothing, _
                                    -dFace1toCenter + dTangentLength1, 0, 0, _
                                    1, 0, 0, _
                                    0, 0, -pipeDiam / 2, _
                                    1, 0, PI)
            oCurvesColl.Add objArcPort1_2
            Set objArcPort1_2 = Nothing
            
            Dim oStPoint As AutoMath.DPosition
            Dim oEnPoint As AutoMath.DPosition
            Dim oCrossSectionPort1 As IngrGeom3D.ComplexString3d
            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set -dFace1toCenter + dTangentLength1, 0, pipeDiam / 2
            Set oCrossSectionPort1 = PlaceTrCString(oStPoint, oCurvesColl)
            Dim objCrossSections As IngrGeom3D.IJElements
            Set objCrossSections = New JObjectCollection
            objCrossSections.Add oCrossSectionPort1
            Set oCurvesColl = Nothing
            Set oCrossSectionPort1 = Nothing
         
            'Forming cross  section at port 2 (splitting the circle into a complex string of two
            'semicircles to facilitate skinning).
             Dim objArcPort2_1 As IngrGeom3D.EllipticalArc3d
            Set objArcPort2_1 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                    Nothing, _
                                    (dFace2toCenter - dTangentLength2) * Cos(parAngle), (dFace2toCenter - dTangentLength2) * Sin(parAngle), 0, _
                                    Cos(parAngle), Sin(parAngle), 0, _
                                    0, 0, pipeDiam2 / 2, _
                                    1, 0, PI)
            Set oCurvesColl = New Collection
            oCurvesColl.Add objArcPort2_1
            Set objArcPort2_1 = Nothing
            
            Dim objArcPort2_2 As IngrGeom3D.EllipticalArc3d
            Set objArcPort2_2 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                    Nothing, _
                                    (dFace2toCenter - dTangentLength2) * Cos(parAngle), (dFace2toCenter - dTangentLength2) * Sin(parAngle), 0, _
                                    Cos(parAngle), Sin(parAngle), 0, _
                                    0, 0, -pipeDiam2 / 2, _
                                    1, 0, PI)
            oCurvesColl.Add objArcPort2_2
            Set objArcPort2_2 = Nothing
            
            Dim oCrossSectionPort2 As IngrGeom3D.ComplexString3d
            oStPoint.Set (dFace2toCenter - dTangentLength2) * Cos(parAngle), (dFace2toCenter - dTangentLength2) * Sin(parAngle), pipeDiam2 / 2
            Set oCrossSectionPort2 = PlaceTrCString(oStPoint, oCurvesColl)
            objCrossSections.Add oCrossSectionPort2
            Set oCurvesColl = Nothing
            Set oCrossSectionPort2 = Nothing
            
            Dim dElbowStarttoCenter As Double
            Dim dElbowEndtoCenter   As Double
            Dim dHoriCompofPipeRadius2 As Double
            Dim dVerCompofPipeRadius2 As Double
            dElbowStarttoCenter = dFace1toCenter - dTangentLength1
            dElbowEndtoCenter = dFace2toCenter - dTangentLength2
            dVerCompofPipeRadius2 = (pipeDiam2 / 2) * Cos(parAngle)
            dHoriCompofPipeRadius2 = (pipeDiam2 / 2) * Sin(parAngle)
            
            'Constructing Trace curves for skinning.
            Dim objEllipticalArc1 As IngrGeom3D.EllipticalArc3d
            Set objEllipticalArc1 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                    Nothing, _
                                    -dElbowStarttoCenter, dElbowEndtoCenter * Sin(parAngle) + dVerCompofPipeRadius2, 0, _
                                    0, 0, 1, _
                                    0, -(dElbowEndtoCenter * Sin(parAngle) + dVerCompofPipeRadius2 - (0.5 * pipeDiam)), 0, _
                                    (dElbowStarttoCenter + dElbowEndtoCenter * Cos(parAngle) - dHoriCompofPipeRadius2) / _
                                    (dElbowEndtoCenter * Sin(parAngle) + dVerCompofPipeRadius2 - (0.5 * pipeDiam)), 0, _
                                     PI / 2)
            
            Dim objEllipticalArc2 As IngrGeom3D.EllipticalArc3d
            Set objEllipticalArc2 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                    Nothing, _
                                    -dElbowStarttoCenter, dElbowEndtoCenter * Sin(parAngle) - dVerCompofPipeRadius2, 0, _
                                     0, 0, 1, _
                                     0, -(dElbowEndtoCenter * Sin(parAngle) - dVerCompofPipeRadius2 + (0.5 * pipeDiam)), 0, _
                                    (dElbowStarttoCenter + dElbowEndtoCenter * Cos(parAngle) + dHoriCompofPipeRadius2) / _
                                    (dElbowEndtoCenter * Sin(parAngle) - dVerCompofPipeRadius2 + (0.5 * pipeDiam)), 0, _
                                     PI / 2)
            
            Dim objTraceCurves As IngrGeom3D.IJElements
            Set objTraceCurves = New JObjectCollection
            objTraceCurves.Add objEllipticalArc1
            objTraceCurves.Add objEllipticalArc2
            
            Dim objBend As IngrGeom3D.IJElements
            Set objBend = geomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
                                                objTraceCurves, objCrossSections, True)

            'Set the output
            Dim objOuput As Object
            Dim iCount As Integer
            For Each objOuput In objBend
                If iOutput = 0 Then
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Elbow", objOuput
                Else
                    iCount = 0
                    m_OutputColl.AddOutput "Elbow" & iCount, objOuput
                    iCount = iCount + 1
                End If
            Next

            'Remove References
            For iCount = 1 To objBend.Count
                objBend.Remove 1
            Next iCount

            Set objOuput = Nothing
            Set objBend = Nothing
            Set objCrossSections = Nothing
            Set objTraceCurves = Nothing
            Set objEllipticalArc1 = Nothing
            Set objEllipticalArc2 = Nothing
            
             'Place Nozzle 1
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            
            oPlacePoint.Set -dFace1toCenter - cptOffset + depth, 0, 0
            oDir.Set -1, 0, 0
        
            If CmpDblLessThan(dTangentLength1, flangeThick) Then dTangentLength1 = flangeThick
        
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                                oPlacePoint, dTangentLength1)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
            
            'Place Nozzle 2
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            
            oPlacePoint.Set (dFace2toCenter + cptOffset2 - depth2) * Cos(parAngle), (dFace2toCenter + cptOffset2 - depth2) * Sin(parAngle), 0
            oDir.Set Cos(parAngle), Sin(parAngle), 0
            If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2
            
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                                oPlacePoint, dTangentLength2)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
    
            If lPartDataBasis >= 550 And lPartDataBasis <= 569 Then
                Dim pipeDiam3       As Double
                Dim flangeThick3    As Double
                Dim flangeDiam3     As Double
                Dim cptOffset3      As Double
                Dim depth3          As Double
                Dim dFace3toCenter  As Double
    
                RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, cptOffset3, depth3
                Dim oHeelInlet                  As Object
                Dim dCentertoCenterHeelInlet    As Double
                Dim dHeelStPos                  As Double
                
                Set oStPoint = New AutoMath.DPosition
                Set oEnPoint = New AutoMath.DPosition
                Select Case lPartDataBasis
                Case 550, 552
                    dFace3toCenter = parHeelInletSeattoCenter - cptOffset3 + depth3
                    dHeelStPos = -0.8 * dFace1toCenter
                Case 551, 553
                    dFace3toCenter = parHeelInletFacetoCenter
                    dHeelStPos = -0.8 * dFace1toCenter
                Case 554, 556, 558, 560
                    dFace3toCenter = (parSeat1toSeatHeelInlet - dSeat1toCenter) - cptOffset3 + depth3
                    dHeelStPos = -0.8 * dFace1toCenter
                Case 555, 557, 559, 561
                    dFace3toCenter = parFace1toFaceHeelInlet - dFace1toCenter
                    dHeelStPos = -0.8 * dFace1toCenter
                Case 562, 564, 566, 568
                    dFace3toCenter = (parSeat1toSeatHeelInlet - dSeat1toCenter) - cptOffset3 + depth3
                    dCentertoCenterHeelInlet = parCentertoCenterHeelInlet
                    dHeelStPos = 0
                Case 563, 565, 567, 569
                    dFace3toCenter = parFace1toFaceHeelInlet - dFace1toCenter
                    dCentertoCenterHeelInlet = parCentertoCenterHeelInlet
                    dHeelStPos = 0
                Case Else
                    GoTo ErrorLabel:
                End Select
                
                oStPoint.Set dHeelStPos, dCentertoCenterHeelInlet, 0
                oEnPoint.Set dFace3toCenter, dCentertoCenterHeelInlet, 0
                Set oHeelInlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
                m_OutputColl.AddOutput "HeelInlet", oHeelInlet
                Set oHeelInlet = Nothing
    
                'Place Nozzle 3
                Set oPlacePoint = New AutoMath.DPosition
                Set oDir = New AutoMath.DVector
                oPlacePoint.Set dFace3toCenter + cptOffset3 - depth3, dCentertoCenterHeelInlet, 0
                oDir.Set 1, 0, 0
                Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
                'Set the output
                m_OutputColl.AddOutput "Nozzle3", objNozzle
                Set objNozzle = Nothing
                Set oPlacePoint = Nothing
                Set oDir = Nothing
                Set oStPoint = Nothing
                Set oEnPoint = Nothing
    
            ElseIf lPartDataBasis >= 570 And lPartDataBasis <= 579 Then
                RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, cptOffset3, depth3
                Dim oSideOpening                As Object
                Dim dOpeningCentertoBendCenter1 As Double
                Dim dOpeningCentertoBendCenter2 As Double
                
                Set oStPoint = New AutoMath.DPosition
                Set oEnPoint = New AutoMath.DPosition
                Select Case lPartDataBasis
                Case 570, 572
                    dFace3toCenter = parSeatSideInlettoCenter - cptOffset3 + depth3
                    dOpeningCentertoBendCenter1 = dSeat1toCenter - parSideInletCentertoSeat
                    dOpeningCentertoBendCenter2 = dSeat2toCenter - parSideInletCentertoSeat
                Case 571, 573
                    dFace3toCenter = parFaceSideInlettoCenter
                    dOpeningCentertoBendCenter1 = dFace1toCenter - parSideInletCentertoFace
                    dOpeningCentertoBendCenter2 = dFace2toCenter - parSideInletCentertoFace
                Case 574, 576, 578
                    dFace3toCenter = parFaceSideOutlettoCenter
                    dOpeningCentertoBendCenter1 = (1 - Sin(PI / 4)) * dElbowRadius
                    dOpeningCentertoBendCenter2 = (1 - Cos(PI / 4)) * dElbowRadius
                Case 575, 577, 579
                    dFace3toCenter = parSeatSideOutlettoCenter - cptOffset3 + depth3
                    dOpeningCentertoBendCenter1 = (1 - Sin(PI / 4)) * dElbowRadius
                    dOpeningCentertoBendCenter2 = (1 - Cos(PI / 4)) * dElbowRadius
                Case Else
                    GoTo ErrorLabel:
                End Select
                
                oStPoint.Set -dOpeningCentertoBendCenter1, dOpeningCentertoBendCenter2, 0
                oEnPoint.Set -dOpeningCentertoBendCenter1, dOpeningCentertoBendCenter2, dFace3toCenter
                Set oSideOpening = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
                m_OutputColl.AddOutput "SideOpening", oSideOpening
                Set oSideOpening = Nothing
    
                'Place Nozzle 3
                Set oPlacePoint = New AutoMath.DPosition
                Set oDir = New AutoMath.DVector
                oPlacePoint.Set -dOpeningCentertoBendCenter1, dOpeningCentertoBendCenter2, (dFace3toCenter + cptOffset3 - depth3)
                oDir.Set 0, 0, 1
                Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
                'Set the output
                m_OutputColl.AddOutput "Nozzle3", objNozzle
                Set objNozzle = Nothing
                Set oPlacePoint = Nothing
                Set oDir = Nothing
                Set oStPoint = Nothing
                Set oEnPoint = Nothing
    
            ElseIf lPartDataBasis >= 580 And lPartDataBasis <= 583 Then
    
                'Create Circular base
                Dim oBase As Object
                Set oStPoint = New AutoMath.DPosition
                Set oEnPoint = New AutoMath.DPosition
                oStPoint.Set 0, -parCentertoBase, 0
                oEnPoint.Set 0, -parCentertoBase + parBaseThickness, 0
    
                Set oBase = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBaseWidth, True)
                m_OutputColl.AddOutput "Base", oBase
                Set oBase = Nothing
                Set oStPoint = Nothing
                Set oEnPoint = Nothing
    
                'Create Supports
                Dim oSupport1 As Object
                Set oStPoint = New AutoMath.DPosition
                Set oEnPoint = New AutoMath.DPosition
    
                oStPoint.Set -parRibThickness / 2, -parCentertoBase, parBaseWidth / 2
                oEnPoint.Set parRibThickness / 2, 0, -parBaseWidth / 2
    
                Set oSupport1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
                m_OutputColl.AddOutput "Support1", oSupport1
    
                Set oSupport1 = Nothing
                Set oStPoint = Nothing
                Set oEnPoint = Nothing
    
                Dim oSupport2 As Object
                Dim lines As Collection
                Dim oLine As Object
                Dim oCString As IngrGeom3D.ComplexString3d
                Dim oVec As AutoMath.DVector
    
                Set oCString = New IngrGeom3D.ComplexString3d
                Set oVec = New AutoMath.DVector
                Set lines = New Collection
                Set oStPoint = New AutoMath.DPosition
                oVec.Set 0, 0, -1
    
                Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, -parBaseWidth / 2, -parCentertoBase, parRibThickness / 2, -parBaseWidth / 2, 0, parRibThickness / 2)
                lines.Add oLine
                Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, -parBaseWidth / 2, 0, parRibThickness / 2, 0, 0, parRibThickness / 2)
                lines.Add oLine
                Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, parRibThickness / 2, 0, 0.9 * dFace2toCenter, parRibThickness / 2)
                lines.Add oLine
                Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0.9 * dFace2toCenter, parRibThickness / 2, parBaseWidth / 2, 0.9 * dFace2toCenter, parRibThickness / 2)
                lines.Add oLine
                Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, parBaseWidth / 2, 0.9 * dFace2toCenter, parRibThickness / 2, parBaseWidth / 2, -parCentertoBase, parRibThickness / 2)
                lines.Add oLine
                Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, parBaseWidth / 2, -parCentertoBase, parRibThickness / 2, -parBaseWidth / 2, -parCentertoBase, parRibThickness / 2)
                lines.Add oLine
    
                oStPoint.Set -parBaseWidth / 2, -parCentertoBase, parRibThickness / 2
                Set oCString = PlaceTrCString(oStPoint, lines)
                Set oSupport2 = PlaceProjection(m_OutputColl, oCString, oVec, parRibThickness, True)
                m_OutputColl.AddOutput "Support2", oSupport2
    
                Set lines = Nothing
                Set oSupport2 = Nothing
                Set oLine = Nothing
                Set oCString = Nothing
                Set oVec = Nothing
            End If
        
        Case Else
            GoTo ErrorLabel:
    End Select

    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
