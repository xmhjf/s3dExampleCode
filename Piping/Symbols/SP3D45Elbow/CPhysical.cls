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
'   Creation Date:  Sunday, Nov 26 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     -----                ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   16.Jan.2006     KKK                 CR-91107  Update elbow symbols to respect Part Data Basis property
'   08.SEP.2006     KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   07-Mar-07       svsmylav            TR-116058: Added new optional input within the part data basis If-Else condition (not to fail the prior to V7 instances)
'   22.Nov.2007     VRG                 CR-127554 Enhanced the existing symbol code to support 1/8 bend, 1/8 bend-extended, 1/8 bend-long cases using PDB values
'   09.Jan.2008     RRK                 CR-134567 Updated elbow symbols to reference PDB values (from 668 to 679) that are specific to 1/8 bends
'   28.Jul.2008     KKC                 CR-145699 Enhance 45 degree and 90 degree elbow symbol to support JIS case
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
    Dim flangeThick1     As Double
    Dim cptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjElbow As Object
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    Dim parFacetoCenter As Double
    Dim parElbowRadius As Double
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
'    parElbowRadius = arrayOfInputs(4)
'    parSeattoCenter = arrayOfInputs(5)
'    parSeat1toCenter = arrayOfInputs(6)
'    parSeat2toCenter = arrayOfInputs(7)
'    parFace1toCenter = arrayOfInputs(8)
'    parFace2toCenter = arrayOfInputs(9)
'    parTangentLength = arrayOfInputs(10)
'    parTangentLength1 = arrayOfInputs(11)
'    parTangentLength2 = arrayOfInputs(12)
'    parAngle = arrayOfInputs(13)
    
    Dim dFixedAngle As Double
    dFixedAngle = PI / 4
    parAngle = dFixedAngle
    
    'Check to see that old instances of the symbol do not fail
    If UBound(arrayOfInputs) >= 13 Then parAngle = arrayOfInputs(13)
    
    iOutput = 0

 ' Insert your code for output 1(Body)
    
    Dim ArcStPointToCenter  As Double
    Dim dElbowRadius As Double
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    'Buttwelded 45 degree elbow the FacetoCenter is the elbow radius, assume Bolted the same.
    'Socket welded 45 degree elbow the FacetoCenter is the face to the origin of the elbow at the intersection of the two legs.
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick1, flangeDiam, cptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick2, flangeDiam, cptOffset2, depth2
                        
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
'   The symbol implements the following Part Data Bases,
' Symmetrical Elbow-
' a) Default or Face-to-center dimension basis-10
' b)1/8 bend, symmetrical, specified by face-to-center-668
' c) Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).
' d)1/8 bend, symmetrical, specified by seat-to-center-670
' e) Elbow, Symmetrical, specified by Seat to Center and Bend Radius-163
' f)1/8 bend, symmetrical, specified by seat to center and bend radius-672
' g) Elbow, Symmetrical, specified by Face to Center and Bend Radius-164
' h)1/8 bend, symmetrical, specified by face to center and bend radius-674
' i) Elbow, Symmetrical, specified Seat to Center, Tangent length and Bend Radius-161
' j)1/8 bend, symmetrical, specified seat to center, tangent length and bend radius-676
' k) Elbow, Symmetrical, specified Face to Center, Tangent length and Bend Radius-162
' l)1/8 bend, symmetrical, specified face to center, tangent length and bend radius-678

' Asymmetrical Elbow-
' a) Asymmetrical Face-to-Center dimension basis-13
' b) 1/8 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2-669
' c) Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)-36
' d) 1/8 bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2-671
' e) Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius-176
' f) 1/8 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius-673
' g) Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius-177
' h) 1/8 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius-675
' i) Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius-179
' j) 1/8 bend, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius-677
' k) Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent 1, Tangent 2, and Bend Radius-180
' l) 1/8 bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius-679
' m) Elbow based on bend radius and leg offset-55

    Select Case lPartDataBasis
    Case Is <= 1, 10, 668 'Default or Face-to-center dimension basis(10)
                          '1/8 bend, symmetrical, specified by face-to-center(668)
        parFacetoCenter = arrayOfInputs(2)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dTangentLength1 = flangeThick1
        dTangentLength2 = flangeThick2
        'Computing bend radius based on face to center and flange thickness
        dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dFixedAngle / 2)
        
    Case 55 'Elbow based on bend radius and leg offset
        parFacetoCenter = arrayOfInputs(2)
        parElbowRadius = arrayOfInputs(4)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dTangentLength1 = dFace1toCenter - parElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dTangentLength1
        dElbowRadius = parElbowRadius
        
    Case 13, 669 'Asymmetrical Face-to-Center dimension basis(13)
                 '1/8 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2(669)
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        
        'Elbow radius is computed based on smaller of the two face to center dimensions
        'Tangent length is computed such that tangent is created at the port for which face to center dimension is greater
        If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
            dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dFixedAngle / 2)
            dTangentLength1 = flangeThick1
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        Else
            dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dFixedAngle / 2)
            dTangentLength2 = flangeThick2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        End If

    Case 35, 670 'Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(35)
                 '1/8 bend, symmetrical, specified by seat-to-center(670)
        parSeattoCenter = arrayOfInputs(5)
        dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
        dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
        dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dFixedAngle / 2)
        
        'The face to centers may be unequal based on the port generic dimensions
        'Elbow radius is computed based on smaller of the two face to center dimensions
        'Tangent length is computed such that tangent is created at the port for which face to center dimension is greater.
        If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
            dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dFixedAngle / 2)
            dTangentLength1 = flangeThick1
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        Else
            dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dFixedAngle / 2)
            dTangentLength2 = flangeThick2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        End If
        
    Case 36, 671 'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(36)
                 '1/8 bend, asymmetrical, specified by seat-to-center 1 and seat-to-center 2(671)
        parSeat1toCenter = arrayOfInputs(6)
        parSeat2toCenter = arrayOfInputs(7)
        dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
        dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
        
        'Elbow radius is computed based on smaller of the two face to center dimensions
        'Tangent length is computed such that tangent is created at the port for which face to center dimension is greater.
        If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
            dElbowRadius = (dFace1toCenter - flangeThick1) / Tan(dFixedAngle / 2)
            dTangentLength1 = flangeThick1
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        Else
            dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dFixedAngle / 2)
            dTangentLength2 = flangeThick2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        End If
        
    Case 161, 676 'Elbow, Symmetrical, specified Seat to Center, Tangent length and Bend Radius(161)
                  '1/8 bend, symmetrical, specified seat to center, tangent length and bend radius(676)
        parElbowRadius = arrayOfInputs(4)
        parSeattoCenter = arrayOfInputs(5)
        parTangentLength = arrayOfInputs(10)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
        dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        
    Case 162, 678 'Elbow, Symmetrical, specified Face to Center, Tangent length and Bend Radius(162)
                  '1/8 bend, symmetrical, specified face to center, tangent length and bend radius(678)
        parFacetoCenter = arrayOfInputs(2)
        parElbowRadius = arrayOfInputs(4)
        parTangentLength = arrayOfInputs(10)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        
    Case 163, 672  'Elbow, Symmetrical, specified by Seat to Center and Bend Radius(163)
                   '1/8 bend, symmetrical, specified by seat to center and bend radius(672)
        parElbowRadius = arrayOfInputs(4)
        parSeattoCenter = arrayOfInputs(5)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
        dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)

    Case 164, 674  'Elbow, Symmetrical, specified by Face to Center and Bend Radius(164)
                   '1/8 bend, symmetrical, specified by face to center and bend radius(674)
        parFacetoCenter = arrayOfInputs(2)
        parElbowRadius = arrayOfInputs(4)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        
    Case 176, 673 'Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius(176)
                  '1/8 bend, asymmetrical, specified by seat to center 1, seat to center 2 and bend radius(673)
        parElbowRadius = arrayOfInputs(4)
        parSeat1toCenter = arrayOfInputs(6)
        parSeat2toCenter = arrayOfInputs(7)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
        dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
                    
    Case 177, 675 'Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius(177)
                  '1/8 bend, asymmetrical, specified by face to center 1, face to center 2 and bend radius(675)
        parElbowRadius = arrayOfInputs(4)
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
    
    Case 179, 677 'Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius(179)
                  '1/8 bend, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius(677)
        parElbowRadius = arrayOfInputs(4)
        parSeat1toCenter = arrayOfInputs(6)
        parSeat2toCenter = arrayOfInputs(7)
        parTangentLength1 = arrayOfInputs(11)
        parTangentLength2 = arrayOfInputs(12)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
        dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        
    Case 180, 679  'Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent 1, Tangent 2, and Bend Radius(180).
                       '1/8 bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius(679)
        parElbowRadius = arrayOfInputs(4)
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        parTangentLength1 = arrayOfInputs(11)
        parTangentLength2 = arrayOfInputs(12)
        dElbowRadius = parElbowRadius
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dFixedAngle / 2)
        dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dFixedAngle / 2)
    
    Case 182 ' Elbow, asymmetrical, specified by tangent 1, face-to-center 2
        parFace2toCenter = arrayOfInputs(9)
        parTangentLength1 = arrayOfInputs(11)
        dFace1toCenter = parFace2toCenter + parTangentLength1
        dElbowRadius = parFace2toCenter / Tan(dFixedAngle / 2)
        dFace2toCenter = parFace2toCenter
        dTangentLength1 = parTangentLength1
        dTangentLength2 = 0
        
    Case Else
        GoTo ErrorLabel 'Invalid specification.
    End Select
        
    'Recomputing face to center dimensions when elbow is trimmable.
    dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
    dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
    ArcStPointToCenter = dFace1toCenter - dTangentLength1
    
    CP.Set -ArcStPointToCenter, dElbowRadius, 0
    CV.Set 0, 0, 1
  
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

' Construct a circle that will be used to create the revolution
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         CP.x, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                    CV, CP, parAngle, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    
'Check to see that minimum value of tangent lengths is flange thickness
    If CmpDblLessThan(dTangentLength1, flangeThick1) Then dTangentLength1 = flangeThick1
    If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2
    
' Place Nozzle 1
        
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim parPorttoOrigin As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    parPorttoOrigin = dFace1toCenter + cptOffset1 - depth1
    oPlacePoint.Set -parPorttoOrigin, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dTangentLength1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    
    parPorttoOrigin = dFace2toCenter + cptOffset2 - depth2
    oPlacePoint.Set parPorttoOrigin * Cos(parAngle), parPorttoOrigin * Sin(parAngle), 0
    
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dTangentLength2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
