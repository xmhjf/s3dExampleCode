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
'
'   CInsulation.cls
'   Author:         MA
'   Creation Date:  Wednesday, Oct 31 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   72 Degree Elbow.
'   This symbol is developed for CR-127553
'   Source: ASTM A74-06, Table 10, One-Fifth Bend
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   09.Jan.2008     RRK     CR-134567 Updated elbow symbols to reference PDB values (from 642 to 649) that are specific to 1/5 bends
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI           As Double
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Const NEGLIGIBLE_VALUE = 0.0001

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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Integer
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim parFacetoCenter As Double
    Dim parSeattoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parTangentLength As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parElbowRadius As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
    Dim dBendAngle72Deg As Double
    dBendAngle72Deg = PI / 2.5
    parAngle = dBendAngle72Deg
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parSeattoCenter = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parFace2toCenter = arrayOfInputs(5)
'    parSeat1toCenter = arrayOfInputs(6)
'    parSeat2toCenter = arrayOfInputs(7)
'    parTangentLength = arrayOfInputs(8)
'    parTangentLength1 = arrayOfInputs(9)
'    parTangentLength2 = arrayOfInputs(10)
'    parElbowRadius = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    parAngle = arrayOfInputs(13)
    
    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                            flangeDiam, cptOffset, depth

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                                            flangeDiam2, cptOffset2, depth2
 
    'Insert your code for output 4(Insulated Body)
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dInsulationRadius As Double
    Dim dElbowRadius As Double
      
    dInsulationRadius = pipeDiam / 2 + parInsulationThickness
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis

        Case Is <= 1, 10, 642 'Face-to-center dimension basis(10)
                              '1/5 bend, symmetrical, specified by face-to-center(642)
            parFacetoCenter = arrayOfInputs(2)
            dTangentLength1 = flangeThick  'zero tangent lengths
            dTangentLength2 = flangeThick2
            dElbowRadius = (parFacetoCenter - flangeThick) / Tan(dBendAngle72Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 13, 645 'Asymmetrical Face-to-Center dimension basis(13)
                     '1/5 bend, asymmetrical, specified by face-to-center 1 and face-to-center 2(645)
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(5)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle72Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle72Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 35, 643 'Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)(35)
                     '1/5 bend, symmetrical, specified by seat-to-center(643)
            parSeattoCenter = arrayOfInputs(3)
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            'The face to centers may be unequal based on the port generic dimensions.
            'Graphics are created with a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle72Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle72Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 36, 644 'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            'This case has a tangent length on the side of the elbow that has greater face to center value.
            'The lesser of the two face to centers determines the elbow radius and tangent length.
            If CmpDblLessThanOrEqualTo(dFace1toCenter, dFace2toCenter) Then
                dElbowRadius = (dFace1toCenter - flangeThick) / Tan(dBendAngle72Deg / 2)
                dTangentLength1 = flangeThick
                dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            Else
                dElbowRadius = (dFace2toCenter - flangeThick2) / Tan(dBendAngle72Deg / 2)
                dTangentLength2 = flangeThick2
                dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            End If
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 161, 646 'Elbow, Symmetrical, specified Seat to Center, Tangent length and Bend Radius(161)
                      '1/5 bend, symmetrical, specified seat to center, tangent length and bend radius(646)
            parSeattoCenter = arrayOfInputs(3)
            parTangentLength = arrayOfInputs(8)
            parElbowRadius = arrayOfInputs(11)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeattoCenter - cptOffset + depth
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 162, 647 'Elbow, Symmetrical, specified Face to Center, Tangent length and Bend Radius(162)
                      '1/5 bend, symmetrical, specified by face to center, tangent length and bend radius(647)
            parFacetoCenter = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(8)
            parElbowRadius = arrayOfInputs(11)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(dBendAngle72Deg / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 179, 648 'Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius(179)
                      '1/5 bend, asymmetrical, specified seat to center 1, seat to center 2, tangent length 1, tangent length 2 and bend radius(648)
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            parElbowRadius = arrayOfInputs(11)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parSeat1toCenter - cptOffset + depth
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(parAngle / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(parAngle / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case 180, 649  'Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent length 1, Tangent length 2, and Bend Radius(180)
                       '1/5 bend, asymmetrical, specified by face-to-center 1, face-to-center 2, tangent length 1, tangent length 2, and bend radius(649)
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(5)
            parTangentLength1 = arrayOfInputs(9)
            parTangentLength2 = arrayOfInputs(10)
            parElbowRadius = arrayOfInputs(11)
            dElbowRadius = parElbowRadius
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(parAngle / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(parAngle / 2)
            'Face to centers are recomputed when elbow is trimable.
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength1
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + dTangentLength2
        
        Case Else
            GoTo ErrorLabel:
            
    End Select
        
    'Check the value of Tangent Length, and if it is equal to zero, then provide a small value to it
    If CmpDblLessThan(dTangentLength1, flangeThick) Then dTangentLength1 = flangeThick
    If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2

    'Construct a circle that will be used to create the revolution
    If CmpDblGreaterthan(dInsulationRadius, dElbowRadius) Then _
                    dInsulationRadius = dElbowRadius - NEGLIGIBLE_VALUE
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim objCircleIns   As IngrGeom3D.Circle3d
    Set objCircleIns = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        -dFace1toCenter + dTangentLength1, 0, 0, _
                        1, 0, 0, _
                        dInsulationRadius)
 
    'Revolve the circle for 72 degree about the Z axis passing by the origin
    Dim CP As AutoMath.DPosition  'revolution center point
    Dim CV As AutoMath.DVector  'rotation vector for rotation
    Set CP = New AutoMath.DPosition
    Set CV = New AutoMath.DVector
    CP.Set -dFace1toCenter + dTangentLength1, dElbowRadius, 0
    CV.Set 0, 0, 1
    Dim ObjInsulatedBody As Object
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleIns, _
                                     CV, CP, parAngle, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set objCircleIns = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set oGeomFactory = Nothing
    
    'Construct Insulation for the tangential portion 1
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    If CmpDblGreaterthan(dTangentLength1, flangeThick + LINEAR_TOLERANCE) Then
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set -dFace1toCenter + dTangentLength1, 0, 0

        Dim ObjInsTangentPortion1 As Object
        Set ObjInsTangentPortion1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam + parInsulationThickness * 2, True)

        'Set the output
        m_OutputColl.AddOutput "InsTanPortion_", ObjInsTangentPortion1
        Set ObjInsTangentPortion1 = Nothing
    End If
 
    'Construct Insulation for the Tangential portion 2
    If CmpDblGreaterthan(dTangentLength2, flangeThick2 + LINEAR_TOLERANCE) Then
        oStPoint.Set dFace2toCenter * Cos(parAngle), _
                                dFace2toCenter * Sin(parAngle), 0
        oEnPoint.Set (dFace2toCenter - dTangentLength2) * Cos(parAngle), _
                            (dFace2toCenter - dTangentLength2) * Sin(parAngle), 0
  
        Dim ObjInsTangentPortion2 As Object
        Set ObjInsTangentPortion2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                pipeDiam2 + parInsulationThickness * 2, True)
 
        'Set the output
        m_OutputColl.AddOutput "InsTanPortion_", ObjInsTangentPortion2
        Set ObjInsTangentPortion2 = Nothing
     End If

    'Place Insulation for Port 1
    If CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
        If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0
  
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        flangeDiam + 2 * parInsulationThickness, True)

        'Set the output
        m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
     End If
 
    'Place Insulation for Port 2
    If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
        If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
        oStPoint.Set dFace2toCenter * Cos(parAngle), dFace2toCenter * Sin(parAngle), 0
        oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness) * Cos(parAngle), _
             (dFace2toCenter - flangeThick2 - parInsulationThickness) * Sin(parAngle), 0

        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                flangeDiam2 + 2 * parInsulationThickness, True)

        'Set the output
        m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
     End If
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
   
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    End Sub
