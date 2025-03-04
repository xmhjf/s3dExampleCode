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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         dkl
'   Creation Date:  Monday, 12, Feb 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details are taken from the Design document, Appendix E-26.
'
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------     -----    ------------------
'   12.Feb.2007     dkl      CR-113390  Created the symbol.
'   7.Nov.2007      VRG      TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                            in which case small cylinders of negligible thickness will be created

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double

    Dim iOutput     As Double

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parFacetoCenter As Double
    Dim parTangentLength As Double
    Dim parBendRadius As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double

    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parTangentLength1 = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    parTangentLength2 = arrayOfInputs(5)
    parFacetoCenter = arrayOfInputs(6)
    parTangentLength = arrayOfInputs(7)
    parBendRadius = arrayOfInputs(8)
    parAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0

    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 160 'TangentLength and BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dElbowRadius = parBendRadius
            dFace1toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength
            dFace2toCenter = dFace1toCenter
        
        Case 165    'FacetoCenter and TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFacetoCenter - parTangentLength) / Tan(parAngle / 2)
        
        Case 170   'TangentLength1, TangentLength2 and BendRadius
            dElbowRadius = parBendRadius
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dFace1toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength1
            dFace2toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength2
            
        Case 175   'Face1toCenter, TangentLength1 and BendRadius
            dElbowRadius = parBendRadius
            dTangentLength1 = parTangentLength1
            dFace1toCenter = parTangentLength1 + parBendRadius * Tan(parAngle / 2)
            dTangentLength2 = 0
            dFace2toCenter = parBendRadius * Tan(parAngle / 2)
            
        Case 180   'Face1toCenter, TangentLength1, Face2toCenter,
                                    'TangentLength2 and BendRadius
            dElbowRadius = parBendRadius
            dFace1toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength1
            dFace2toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength2
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            
        Case 185   'Face1toCenter, TangentLength1 and TangentLength2
            dFace1toCenter = parFace1toCenter
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFace1toCenter - parTangentLength1) / Tan(parAngle / 2)
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + parTangentLength2
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    End Select
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                    sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                        sptOffset, depth
     
' Insert your code for output 4(Insulated Tangent At Port1)
    Dim objInsulatedTangent1 As Object
    Dim dArcStPointToCenter  As Double
    Dim dInsulationRadius As Double
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    If CmpDblEqual(dTangentLength1, flangeThick) Then
        dTangentLength1 = dTangentLength1 + NEGLIGIBLE_THICKNESS
    End If
    dArcStPointToCenter = dFace1toCenter - dTangentLength1
    dInsulationRadius = pipeDiam / 2 + parInsulationThickness
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set -dFace1toCenter + flangeThick, 0, 0
    Set oEndPoint = New AutoMath.DPosition
    oEndPoint.Set -dArcStPointToCenter, 0, 0
    
    Set objInsulatedTangent1 = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                2 * dInsulationRadius, True)
                                
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedTangent1
    Set objInsulatedTangent1 = Nothing

' Insert your code for output 5(Insulated Body)
    dInsulationRadius = pipeDiam / 2 + parInsulationThickness
    'Ensure that the hub radius is not greater than the distance from the center
    'of rotation to the elbow center line
    If CmpDblGreaterthan(dInsulationRadius, dElbowRadius) Then
        dInsulationRadius = dElbowRadius
    End If

    'Construct a circle that will be used to create the revolution
    Dim objCircleI   As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -dArcStPointToCenter, 0, 0, 1, 0, 0, dInsulationRadius)
                         
    'Revolve the circle about the Z axis passing by the origin
    Dim oAxisVect As AutoMath.DVector   'vector specifying axis of rotation
    Dim oCenPoint As AutoMath.DPosition 'arc center point
    Dim objInsulatedBody As Object
    
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set -dArcStPointToCenter, dElbowRadius, 0
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 0, 1
    
    Set objInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    oAxisVect, oCenPoint, parAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
    Set objInsulatedBody = Nothing
    Set geomFactory = Nothing
    Set oCenPoint = Nothing
    Set oAxisVect = Nothing
    Set objCircleI = Nothing

' Insert your code for output 6(Insulated Tangent At Port2)
    Dim objInsulatedTangent2 As Object
    If CmpDblEqual(dTangentLength2, flangeThick2) Then
        dTangentLength2 = dTangentLength2 + NEGLIGIBLE_THICKNESS
    End If
    dArcStPointToCenter = dFace2toCenter - dTangentLength2
    dInsulationRadius = pipeDiam2 / 2 + parInsulationThickness
    oStPoint.Set (dFace2toCenter - flangeThick2) * Cos(parAngle), _
                (dFace2toCenter - flangeThick2) * Sin(parAngle), 0
    oEndPoint.Set dArcStPointToCenter * Cos(parAngle), _
                 dArcStPointToCenter * Sin(parAngle), 0
    
    Set objInsulatedTangent2 = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                2 * dInsulationRadius, True)
                                
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedTangent2
    Set objInsulatedTangent2 = Nothing

' Insert your code for output 7(Insulated Port1)
    Dim objInsulatedPort1 As Object
    Dim dInsDiameterPort  As Double
    
    oStPoint.Set -dFace1toCenter, 0, 0
    oEndPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsDiameterPort = flangeDiam + 2 * parInsulationThickness
    If CmpDblEqual(flangeThick, 0) Then
        oEndPoint.Set -dFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
    End If
    Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                                                dInsDiameterPort, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort1
    Set objInsulatedPort1 = Nothing

' Insert your code for output 8(Insulated Port2)
    Dim objInsulatedPort2 As Object
    
    oStPoint.Set dFace2toCenter * Cos(parAngle), dFace2toCenter * Sin(parAngle), 0
    oEndPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness) * Cos(parAngle), _
                (dFace2toCenter - flangeThick2 - parInsulationThickness) * Sin(parAngle), 0
    dInsDiameterPort = pipeDiam2 + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then dInsDiameterPort = flangeDiam2 + 2 * parInsulationThickness
    If CmpDblEqual(flangeThick2, 0) Then
        oEndPoint.Set (dFace2toCenter - NEGLIGIBLE_THICKNESS) * Cos(parAngle), (dFace2toCenter - NEGLIGIBLE_THICKNESS) * Sin(parAngle), 0
    End If
    Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                                                dInsDiameterPort, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedPort2
    Set objInsulatedPort2 = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
