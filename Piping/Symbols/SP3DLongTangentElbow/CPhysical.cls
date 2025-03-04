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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: dkl
'   Creation Date:  Monday, 12, Feb 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details are taken from the Design document, Appendix E-26.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   12.Feb.2007     dkl  CR-113390  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
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

    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parFacetoCenter As Double
    Dim parTangentLength As Double
    Dim parBendRadius As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double
    
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
    
' This symbol is based on the following Part data basis values that govern its geometry,
' Symmetrical Elbow-
'   (Part data Basis value -160): TangentLength and BendRadius
'   (Part data Basis value -165): FacetoCenter and TangentLength
' Unsymmetrical Elbow-
'   (Part data Basis value -170): TangentLength1, TangentLength2 and BendRadius
'   (Part data Basis value -175): Face1toCenter, TangentLength1 and BendRadius
'   (Part data Basis value -180): Face1toCenter, TangentLength1, Face2toCenter,
'                                           TangentLength2 and BendRadius
'   (Part data Basis value -185): Face1toCenter, TangentLength1 and TangentLength2
    
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
    
' Insert your code for output 1(Elbow)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                flangeDiam, sptOffset, depth
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim dArcStPointToCenter  As Double
    
    dArcStPointToCenter = dFace1toCenter - dTangentLength1
    ' Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -dArcStPointToCenter, 0, 0, 1, 0, 0, pipeDiam / 2)
    ' Revolve the circle about the Z axis
    Dim ObjElbow As Object
    Dim oAxisVect As AutoMath.DVector   'vector specifying axis of rotation
    Dim oCenPoint As AutoMath.DPosition 'arc center point
    
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set -dArcStPointToCenter, dElbowRadius, 0
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 0, 1
    
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, oAxisVect, oCenPoint, _
                                                        parAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set oCenPoint = Nothing
    Set oAxisVect = Nothing
    Set objCircle = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1
    Dim dNozzleLength As Double
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim dPorttoOrigin As Double
    
    dPorttoOrigin = -dFace1toCenter - sptOffset + depth
    Set oPlacePoint = New AutoMath.DPosition
    oPlacePoint.Set dPorttoOrigin, 0, 0
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0

    dNozzleLength = dTangentLength1
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                    sptOffset, depth
    dPorttoOrigin = dFace2toCenter + sptOffset - depth
    oPlacePoint.Set dPorttoOrigin * Cos(parAngle), dPorttoOrigin * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0

    dNozzleLength = dTangentLength2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

