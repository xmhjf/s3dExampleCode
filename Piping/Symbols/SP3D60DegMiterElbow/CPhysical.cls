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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         Nabil Nicolas
'   Creation Date:  Friday, Oct 20 2006
'   Description:    60 degree Miter Elbow based on code of 45 degree Miter Elbow by svsmylav
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim pipeDiam2        As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoCenter As Double
    Dim parNoOfMiterCuts As Long
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parNoOfMiterCuts = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

'   Use parameter to store elbow angle
    Dim dElbowAng As Double
    dElbowAng = PI / 3

' Insert your code for output 3(Tangent at Port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim dElbowRadius As Double

    dElbowRadius = parFacetoCenter / Tan(dElbowAng / 2)
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parFacetoCenter, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)

'   Determine the angle at which the ellipse need to be created
    Dim dAngle As Double
    
    dAngle = (dElbowAng) / (2 * parNoOfMiterCuts)

    Dim ellNormalX As Double
    Dim ellNormalY As Double
    Dim ellNormalZ As Double
    ellNormalX = Cos(dAngle)
    ellNormalY = Sin(dAngle)
    ellNormalZ = 0
    Dim ell1CenterX As Double
    Dim ell1CenterY As Double
    Dim ell1CenterZ As Double
    ell1CenterX = -parFacetoCenter + dElbowRadius * Tan(dAngle)
    ell1CenterY = 0
    ell1CenterZ = 0

    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    MajorX = -Tan(dAngle) * pipeDiam / 2
    MajorY = pipeDiam / 2
    MajorZ = 0
    
    Dim ObjEllipse As Object
    Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ell1CenterX, ell1CenterY, ell1CenterZ, _
                                        ellNormalX, ellNormalY, ellNormalZ, _
                                        MajorX, MajorY, MajorZ, Cos(dAngle))
'   Created Ruled surface
    Dim ObjTangentAtPort1 As Object
    Set ObjTangentAtPort1 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle, ObjEllipse, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangentAtPort1
    Set ObjTangentAtPort1 = Nothing
    
' Insert your code for output 4(Tangent at Port 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2

    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parFacetoCenter * Cos(dElbowAng), _
                    parFacetoCenter * Sin(dElbowAng), 0, _
                        Cos(dElbowAng), Sin(dElbowAng), 0, _
                        pipeDiam2 / 2)
    

    ell1CenterX = parFacetoCenter * Cos(dElbowAng) _
            - dElbowRadius * Tan(dAngle) * Cos(dElbowAng)
    ell1CenterY = parFacetoCenter * Sin(dElbowAng) _
            - dElbowRadius * Tan(dAngle) * Sin(dElbowAng)
    ell1CenterZ = 0
    
    Dim dCurrentSegAngle As Double
    Dim dCosComp As Double
    Dim dSinComp As Double
    dCurrentSegAngle = dElbowAng - dAngle
    
    ellNormalX = Cos(dCurrentSegAngle)
    ellNormalY = Sin(dCurrentSegAngle)
    ellNormalZ = 0
    
    dCosComp = (0.5 * pipeDiam / Cos(dAngle)) * Cos(dCurrentSegAngle)
    dSinComp = (0.5 * pipeDiam / Cos(dAngle)) * Sin(dCurrentSegAngle)
            
    MajorX = dSinComp
    MajorY = -dCosComp
    MajorZ = 0

    Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ell1CenterX, ell1CenterY, ell1CenterZ, _
                                        ellNormalX, ellNormalY, ellNormalZ, _
                                        MajorX, MajorY, MajorZ, Cos(dAngle))

'   Created Ruled surface
    Dim ObjTangentAtPort2 As Object
    Set ObjTangentAtPort2 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle, ObjEllipse, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangentAtPort2
    Set ObjTangentAtPort2 = Nothing

' Insert your code for output (Sections other than tangents)
'   Note: These are variable outputs and are applicable only if Number Of Miter Cuts is greater than one
    If parNoOfMiterCuts > 1 Then
        Dim iCount As Double
        Dim ObjEllipse2 As Object
        Dim lNumMiddleSegments As Long
        Dim dMiddleSegAngle As Double
        lNumMiddleSegments = (parNoOfMiterCuts - 1)
        dMiddleSegAngle = (dElbowAng - 2 * dAngle) / lNumMiddleSegments
        dCurrentSegAngle = dElbowAng - dAngle - dMiddleSegAngle
        Dim dMidOfCurrentSegAngle As Double
        Dim ObjMiddleSegment As Object
        dMidOfCurrentSegAngle = dCurrentSegAngle + dMiddleSegAngle / 2
        
        For iCount = 1 To lNumMiddleSegments
            ell1CenterX = -parFacetoCenter + dElbowRadius * Sin(dMidOfCurrentSegAngle) _
                    - dElbowRadius * Tan(dMiddleSegAngle / 2) * Cos(dMidOfCurrentSegAngle)
            ell1CenterY = dElbowRadius * (1 - Cos(dMidOfCurrentSegAngle)) _
                    - dElbowRadius * Tan(dMiddleSegAngle / 2) * Sin(dMidOfCurrentSegAngle)
            ell1CenterZ = 0

            ellNormalX = -Cos(dCurrentSegAngle)
            ellNormalY = -Sin(dCurrentSegAngle)
            ellNormalZ = 0

            dCosComp = (0.5 * pipeDiam / Cos(dMiddleSegAngle / 2)) * Cos(dCurrentSegAngle)
            dSinComp = (0.5 * pipeDiam / Cos(dMiddleSegAngle / 2)) * Sin(dCurrentSegAngle)
            MajorX = dSinComp
            MajorY = -dCosComp
            MajorZ = 0
            
            Set ObjEllipse2 = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                ell1CenterX, ell1CenterY, ell1CenterZ, _
                                                ellNormalX, ellNormalY, ellNormalZ, _
                                                MajorX, MajorY, MajorZ, Cos(dAngle))
            Set ObjMiddleSegment = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            ObjEllipse, ObjEllipse2, True)

'           Set the output
            m_OutputColl.AddOutput "MiddleSegment_", ObjMiddleSegment
'           Check to see if next section need to be placed
            dCurrentSegAngle = dCurrentSegAngle - dMiddleSegAngle
            dMidOfCurrentSegAngle = dCurrentSegAngle + dMiddleSegAngle / 2
            Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                ell1CenterX, ell1CenterY, ell1CenterZ, _
                                                ellNormalX, ellNormalY, ellNormalZ, _
                                                MajorX, MajorY, MajorZ, Cos(dAngle))
        Next iCount
    End If
    Set geomFactory = Nothing
    Set objCircle = Nothing
    Set ObjEllipse = Nothing
    Set ObjEllipse2 = Nothing
    Set ObjMiddleSegment = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    Dim dPorttoOrigin As Double
    dPorttoOrigin = parFacetoCenter + sptOffset2 - depth2
    oPlacePoint.Set dPorttoOrigin * Cos(dElbowAng), dPorttoOrigin * Sin(dElbowAng), 0
    oDir.Set Cos(dElbowAng), Sin(dElbowAng), 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
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
