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
'   Copyright (c) 2006-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Jun 20 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   08.SEP.2006     KKC        DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.Aug.2007     VRG        TR-124960 Insulation diameter for ports was computed based on End preparation
'   7.Nov.2007      VRG        TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                              in which case small cylinders of negligible thickness will be created
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

    Dim ObjInsulatedNozzle1 As Object
    Dim ObjInsulatedNozzle2 As Object
    Dim iOutput     As Integer
    Dim parFacetoCenter As Double
    Dim parNoOfMiterCuts As Long
    Dim parInsulationThickness As Double
    
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parNoOfMiterCuts = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
    
'   Use parameter to store elbow angle
    Dim dElbowAng As Double
    dElbowAng = PI / 2
    
' Insert your code for output 1(Tangent at Port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim dElbowRadius As Double

    dElbowRadius = parFacetoCenter
'   Ensure that the InsulationThickness does not exceed the difference of ElbowRadius and PipeRadius
    If CmpDblGreaterthanOrEqualTo(parInsulationThickness, dElbowRadius - pipeDiam / 2) Then _
        parInsulationThickness = dElbowRadius - pipeDiam / 2 - NEGLIGIBLE_THICKNESS
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -dElbowRadius, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2 + parInsulationThickness)

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
    ell1CenterX = -dElbowRadius * (1 - Tan(dAngle))

    ell1CenterY = 0
    ell1CenterZ = 0
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    MajorX = -Tan(dAngle) * (pipeDiam / 2 + parInsulationThickness)
    MajorY = pipeDiam / 2 + parInsulationThickness
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

' Insert your code for output 2(Tangent at Port 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'   Ensure that the InsulationThickness does not exceed the difference of ElbowRadius and PipeRadius
    If CmpDblGreaterthanOrEqualTo(parInsulationThickness, dElbowRadius - pipeDiam / 2) Then _
        parInsulationThickness = dElbowRadius - pipeDiam / 2 - NEGLIGIBLE_THICKNESS
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, dElbowRadius, 0, _
                        0, 1, 0, _
                        pipeDiam / 2 + parInsulationThickness)

    ell1CenterX = 0
    ell1CenterY = dElbowRadius * (1 - Tan(dAngle))
    ell1CenterZ = 0

    ellNormalX = Cos(dElbowAng - dAngle)
    ellNormalY = Sin(dElbowAng - dAngle)
    ellNormalZ = 0

    MajorX = pipeDiam / 2 + parInsulationThickness
    MajorY = -Tan(dAngle) * (pipeDiam / 2 + parInsulationThickness)
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
        Dim dCurrentSegAngle As Double
        Dim dMiddleSegAngle As Double
        lNumMiddleSegments = (parNoOfMiterCuts - 1)
        dMiddleSegAngle = (dElbowAng - 2 * dAngle) / lNumMiddleSegments
        dCurrentSegAngle = dElbowAng - dAngle - dMiddleSegAngle
        Dim dMidOfCurrentSegAngle As Double
        Dim ObjMiddleSegment As Object
        dMidOfCurrentSegAngle = dCurrentSegAngle + dMiddleSegAngle / 2
        Dim dCosComp As Double
        Dim dSinComp As Double

        For iCount = 1 To lNumMiddleSegments
            ell1CenterX = -dElbowRadius * (1 - Sin(dMidOfCurrentSegAngle)) _
                    - dElbowRadius * Tan(dMiddleSegAngle / 2) * Cos(dMidOfCurrentSegAngle)
            ell1CenterY = dElbowRadius * (1 - Cos(dMidOfCurrentSegAngle)) _
                    - dElbowRadius * Tan(dMiddleSegAngle / 2) * Sin(dMidOfCurrentSegAngle)
            ell1CenterZ = 0

            ellNormalX = -Cos(dCurrentSegAngle)
            ellNormalY = -Sin(dCurrentSegAngle)
            ellNormalZ = 0

            dCosComp = ((0.5 * pipeDiam + parInsulationThickness) / Cos(dMiddleSegAngle / 2)) * Cos(dCurrentSegAngle)
            dSinComp = ((0.5 * pipeDiam + parInsulationThickness) / Cos(dMiddleSegAngle / 2)) * Sin(dCurrentSegAngle)
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


 ' Insert your code for output 7(Insulated Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim parInsulationDiameter As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    'Providing a check on end point for non flanged ends in which case a small cylinder of negligible thickness is created
    stPoint.Set -parFacetoCenter, 0, 0
    If CmpDblGreaterthanOrEqualTo(-parFacetoCenter + flangeThick + parInsulationThickness, 0) Then
        enPoint.Set 0, 0, 0
    Else
        enPoint.Set -parFacetoCenter + flangeThick + parInsulationThickness, 0, 0
    End If
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set -parFacetoCenter + NEGLIGIBLE_THICKNESS, 0, 0
    End If
    
    Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    m_OutputColl.AddOutput "InsulatedNozzle1", ObjInsulatedNozzle1
    Set ObjInsulatedNozzle1 = Nothing

 ' Insert your code for output 8(Insulated Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    'Providing a check on end point for non flanged ends in which case a small cylinder of negligible thickness is created
    stPoint.Set parFacetoCenter * Cos(dElbowAng), parFacetoCenter * Sin(dElbowAng), 0
    If CmpDblLessThanOrEqualTo(parFacetoCenter - flangeThick - parInsulationThickness, 0) Then
        enPoint.Set 0, 0, 0
    Else
        enPoint.Set (parFacetoCenter - (flangeThick + parInsulationThickness)) * Cos(dElbowAng), (parFacetoCenter - (flangeThick + parInsulationThickness)) * Sin(dElbowAng), 0
    End If

    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
         parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
         parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set (parFacetoCenter - NEGLIGIBLE_THICKNESS) * Cos(dElbowAng), (parFacetoCenter - NEGLIGIBLE_THICKNESS) * Sin(dElbowAng), 0
    End If

    Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    m_OutputColl.AddOutput "InsulatedNozzle2", ObjInsulatedNozzle2
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set ObjInsulatedNozzle2 = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
