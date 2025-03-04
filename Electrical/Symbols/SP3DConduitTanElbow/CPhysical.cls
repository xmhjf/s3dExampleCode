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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          kkk
'   Creation Date:  Thursday, December 1 2005
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is based on the following Part data basis values that govern its geometry,
' Symmetrical Elbow-
'   (Part data Basis value -9006): Elbow, symmetrical, specified by TangentLength and BendRadius (implementation similar to Piping part data basis 160).
'   (Part data Basis value -9007): Elbow, symmetrical, specified by seat-to-center, tangent, and bend radius (implementation similar to Piping part data basis 161).
'   (Part data Basis value -9008): Elbow, symmetrical, specified by face-to-center, tangent, and bend radius (implementation similar to Piping part data basis 162).
'   (Part data Basis value -9011): Elbow, symmetrical, specified by FacetoCenter and TangentLength (implementation similar to Piping part data basis 165).
'   (Part data Basis value -9021): Conduit elbow, symmetrical, specified by face-to-face along port 1, tangent length and bend radius (Source:  Cal Conduit Products (http://coatedconduit.com//elbows.htm)).
' Asymmetrical Elbow-
'   (Part data Basis value -9012): Elbow, asymmetrical, specified by TangentLength1, TangentLength2 and BendRadius (implementation similar to Piping part data basis 170).
'   (Part data Basis value -9013): Elbow, asymmetrical, specified by Face1toCenter, TangentLength1 and BendRadius (implementation similar to Piping part data basis 175).
'   (Part data Basis value -9016): Elbow, asymmetrical, specified by seat-to-center 1, seat-to-center 2, tangent 1,
'                                  tangent 2, and bend radius (implementation similar to Piping part data basis 179).
'   (Part data Basis value -9017): Elbow, asymmetrical, specified by Face1toCenter, TangentLength1, Face2toCenter,
'                                           TangentLength2 and BendRadius (implementation similar to Piping part data basis 180).
'   (Part data Basis value -9019): Elbow, asymmetrical, specified by Face1toCenter, TangentLength1 and TangentLength2 (implementation similar to Piping part data basis 185).
'   (Part data Basis value -9022): Conduit elbow, asymmetrical, specified by face-to-face along port 1, tangent length 1, tangent length 2 and bend radius (Source:  Cal Conduit Products (http://coatedconduit.com//elbows.htm)).
'
'   Change History:
'   dd.mmm.yyyy         who                     change description
'   -----------        -----                   ------------------'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  19.May.2007     dkl     CR-89059  Implemented part data basis 9006, 9007, 9008, 9012, 9013, 9016, 9017, 9019, 9021 and 9022.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Const NEGLIGIBLE_VALUE = 0.001

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitOD        As Double
    Dim flangeThick     As Double
    Dim cptOffset1       As Double
    Dim cptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Integer
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parFacetoCenter As Double
    Dim parTangentLength As Double
    Dim parBendRadius As Double
    Dim parAngle As Double
    Dim parSeattoCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parFacetoFacealongPort1 As Double
    
    Dim ObjElbow    As Object
       
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parTangentLength = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parTangentLength1 = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)
'    parTangentLength2 = arrayOfInputs(7)
'    parSeattoCenter = arrayOfInputs(8)
'    parSeat1toCenter = arrayOfInputs(9)
'    parSeat2toCenter = arrayOfInputs(10)
'    parFacetoFacealongPort1 = arrayOfInputs(11)
'    parAngle = arrayOfInputs(12)

    'Retrieve angle for newer insatnces, the older insyances continue with angle as 90 degrees.
    If UBound(arrayOfInputs) <= 3 Then
        parAngle = PI / 2
    Else
        parAngle = arrayOfInputs(12)
    End If

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
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitOD, flangeThick, _
                                                    flangeDiam, cptOffset1, depth1
    Dim ConduitOD1 As Double
    ConduitOD1 = BodyOD(flangeDiam, ConduitOD)
    RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, _
                                                                    cptOffset2, depth2
    Dim ConduitOD2 As Double
    ConduitOD2 = BodyOD(flangeDiam, ConduitOD)
    Dim ConduitBodyOD As Double
    ConduitBodyOD = BodyOD(ConduitOD1, ConduitOD2)
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 9011
            parFacetoCenter = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(3)
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFacetoCenter - parTangentLength) / Tan(parAngle / 2)
    
        Case 9006
            parTangentLength = arrayOfInputs(3)
            parBendRadius = oPipeComponent.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dElbowRadius = parBendRadius
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + parTangentLength
            dFace2toCenter = dFace1toCenter
            
        Case 9007
            parSeattoCenter = arrayOfInputs(8)
            parTangentLength = arrayOfInputs(3)
            parBendRadius = oPipeComponent.BendRadius
            dElbowRadius = parBendRadius
            dFace1toCenter = parSeattoCenter - cptOffset1 + depth1
            dFace2toCenter = parSeattoCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(parAngle / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(parAngle / 2)
            
        Case 9008
            parFacetoCenter = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(3)
            parBendRadius = oPipeComponent.BendRadius
            dElbowRadius = parBendRadius
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(parAngle / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(parAngle / 2)
               
        Case 9012
            parTangentLength1 = arrayOfInputs(5)
            parTangentLength2 = arrayOfInputs(7)
            parBendRadius = oPipeComponent.BendRadius
            dElbowRadius = parBendRadius
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dFace1toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength1
            dFace2toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength2
            
        Case 9013
            parFace1toCenter = arrayOfInputs(4)
            parTangentLength1 = arrayOfInputs(5)
            parBendRadius = oPipeComponent.BendRadius
            dElbowRadius = parBendRadius
            dTangentLength1 = parTangentLength1
            dFace1toCenter = parTangentLength1 + parBendRadius * Tan(parAngle / 2)
            dTangentLength2 = 0
            dFace2toCenter = parBendRadius * Tan(parAngle / 2)

        Case 9016
            parSeat1toCenter = arrayOfInputs(9)
            parSeat2toCenter = arrayOfInputs(10)
            parTangentLength1 = arrayOfInputs(5)
            parTangentLength2 = arrayOfInputs(7)
            parBendRadius = oPipeComponent.BendRadius
            dElbowRadius = parBendRadius
            dFace1toCenter = parSeat1toCenter - cptOffset1 + depth1
            dFace2toCenter = parSeat2toCenter - cptOffset2 + depth2
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(parAngle / 2)
            dTangentLength2 = dFace2toCenter - dElbowRadius * Tan(parAngle / 2)
            
        Case 9017
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(6)
            parTangentLength1 = arrayOfInputs(5)
            parTangentLength2 = arrayOfInputs(7)
            parBendRadius = oPipeComponent.BendRadius
            dElbowRadius = parBendRadius
            dFace1toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength1
            dFace2toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength2
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            
        Case 9019
            parFace1toCenter = arrayOfInputs(4)
            parTangentLength1 = arrayOfInputs(5)
            parTangentLength2 = arrayOfInputs(7)
            dFace1toCenter = parFace1toCenter
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFace1toCenter - parTangentLength1) / Tan(parAngle / 2)
            dFace2toCenter = dElbowRadius * Tan(parAngle / 2) + parTangentLength2
            
        Case 9021
            parTangentLength = arrayOfInputs(3)
            parFacetoFacealongPort1 = arrayOfInputs(11)
            parBendRadius = oPipeComponent.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dElbowRadius = parBendRadius
            dFace1toCenter = dElbowRadius * Tan(parAngle / 2) + parTangentLength
            dFace2toCenter = dFace1toCenter
            
        Case 9022
            parTangentLength1 = arrayOfInputs(5)
            parTangentLength2 = arrayOfInputs(7)
            parFacetoFacealongPort1 = arrayOfInputs(11)
            parBendRadius = oPipeComponent.BendRadius
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dElbowRadius = parBendRadius
            dFace1toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength1
            dFace2toCenter = parBendRadius * Tan(parAngle / 2) + parTangentLength2
            
        
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    
    End Select
    
' Insert your code for output 1(Elbow)
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim dArcStPointToCenter  As Double
    
    dArcStPointToCenter = dFace1toCenter - dTangentLength1
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -dArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                        ConduitBodyOD / 2)
    Set geomFactory = Nothing
'   Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector   'rotation vector for rotation
    CP.Set -dArcStPointToCenter, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, CV, CP, parAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing

' Place Tangent Portion-1
 
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim ObjTangent1 As Object
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Prevent symbol failure for zero tangent length.
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = NEGLIGIBLE_VALUE
    
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set -dFace1toCenter + dTangentLength1, 0, 0
    
    Set ObjTangent1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, ConduitBodyOD, True)
    
    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangent1
    Set ObjTangent1 = Nothing

    
' Place Nozzle 1

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object
    
    oPlacePoint.Set -dFace1toCenter - cptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    Set ObjNozzle1 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
                  
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    
    
' Place Tangent Portion-2
    Dim ObjTangent2 As Object
    
    'Prevent symbol failure for zero tangent length.
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = NEGLIGIBLE_VALUE
    
    oStPoint.Set dFace2toCenter * Cos(parAngle), dFace2toCenter * Sin(parAngle), 0
    oEnPoint.Set (dFace2toCenter - dTangentLength2) * Cos(parAngle), _
                        (dFace2toCenter - dTangentLength2) * Sin(parAngle), 0
    
    Set ObjTangent2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, ConduitBodyOD, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangent2
    Set ObjTangent2 = Nothing

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set (dFace2toCenter + cptOffset2 - depth2) * Cos(parAngle), _
                            (dFace2toCenter + cptOffset2 - depth2) * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    
    Set ObjNozzle2 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 2)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

