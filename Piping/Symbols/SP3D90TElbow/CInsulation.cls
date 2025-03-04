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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Tuesday, Nov 28 2000
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                    change description
'   -----------     -----                  ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006      KKC                    DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.Aug.2007      VRG                    TR-124960 Insulation diameter was computed based on End preparation
'   7.Nov.2007       VRK                    TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                               in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI           As Double
Private RAD          As Double
Private gscadElem    As IJDObject
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    Dim ObjInsulatedNozzle1 As Object
    Dim ObjInsulatedNozzle2 As Object
    Dim iOutput     As Integer
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoCenter As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parAngle = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)
    If parAngle > PI / 2 Then
    parAngle = PI / 2
    End If
    Dim parElbowRadius As Double
    parElbowRadius = parFacetoCenter
    Dim parFacetoCenterofTrimmedE As Double
    parFacetoCenterofTrimmedE = parElbowRadius * Tan(parAngle / 2)

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth

    Dim parInsulationRadius As Double
    parInsulationRadius = pipeDiam / 2 + parInsulationThickness
    
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -parFacetoCenterofTrimmedE, parElbowRadius, 0
    CV.Set 0, 0, 1

' Limit Insulation until parInsulationRadius coincides with the Radius of the Elbow
    If parInsulationRadius <= parElbowRadius Then
  
' Construct a circle that will be used to create the revolution
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -parFacetoCenterofTrimmedE, 0, 0, _
                        1, 0, 0, _
                        parInsulationRadius)
' Create the revolution of circle about the Elbow center of bending

    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircle, _
                    CV, CP, parAngle, True)
'Delete the circle that was created for the revolution
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove

    Else
'Construct a Bspline limited by the Radius of the elbow
        Dim HalfCord As Double
        HalfCord = Sqr(parInsulationRadius * parInsulationRadius - parElbowRadius * parElbowRadius)

        Dim points(1 To 13)   As New AutoMath.DPosition
        Dim objBsplineI  As IngrGeom3D.BSplineCurve3d
        points(1).Set -parFacetoCenterofTrimmedE, parElbowRadius, 0
        points(2).Set -parFacetoCenterofTrimmedE, parElbowRadius, HalfCord
        points(3).Set -parFacetoCenterofTrimmedE, parElbowRadius / 2, parInsulationRadius
        points(4).Set -parFacetoCenterofTrimmedE, 0, parInsulationRadius
        points(5).Set -parFacetoCenterofTrimmedE, -parInsulationRadius / 2, parInsulationRadius
        points(6).Set -parFacetoCenterofTrimmedE, -parInsulationRadius, parInsulationRadius / 2
        points(7).Set -parFacetoCenterofTrimmedE, -parInsulationRadius, 0
        points(8).Set -parFacetoCenterofTrimmedE, -parInsulationRadius, -parInsulationRadius / 2
        points(9).Set -parFacetoCenterofTrimmedE, -parInsulationRadius / 2, -parInsulationRadius
        points(10).Set -parFacetoCenterofTrimmedE, 0, -parInsulationRadius
        points(11).Set -parFacetoCenterofTrimmedE, parElbowRadius / 2, -parInsulationRadius
        points(12).Set -parFacetoCenterofTrimmedE, parElbowRadius, -HalfCord
        points(13).Set -parFacetoCenterofTrimmedE, parElbowRadius, 0


        Set objBsplineI = PlaceTrBspline(5, points)
        ' Revolve the Bspline for 90 degree about the Z axis passing by the origin
    
    
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objBsplineI, _
                    CV, CP, parAngle, True)
'Delete the Bsplie that was created for the revolution
    Set gscadElem = objBsplineI
    Set objBsplineI = Nothing
    gscadElem.Remove

    End If

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
' Insert your code for output 5(Insulated Nozzle1)
   RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parInsulationDiameter As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -parFacetoCenterofTrimmedE, 0, 0
    If CmpDblGreaterthanOrEqualTo(-parFacetoCenterofTrimmedE + flangeThick + parInsulationThickness, 0) Then
        enPoint.Set 0, 0, 0
    Else
        enPoint.Set -parFacetoCenterofTrimmedE + flangeThick + parInsulationThickness, 0, 0
    End If
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
         parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
         parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set -parFacetoCenterofTrimmedE + NEGLIGIBLE_THICKNESS, 0, 0
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
       
' Set the output
    m_OutputColl.AddOutput "InsNozzle1", ObjInsulatedNozzle1
    Set ObjInsulatedNozzle1 = Nothing
    
' Insert your code for output 6(Insulated Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    stPoint.Set parFacetoCenterofTrimmedE * Cos(parAngle), parFacetoCenterofTrimmedE * Sin(parAngle), 0
    If CmpDblLessThanOrEqualTo((parFacetoCenterofTrimmedE - (flangeThick + parInsulationThickness)), 0) Then
        enPoint.Set 0, 0, 0
    Else
        enPoint.Set (parFacetoCenterofTrimmedE - (flangeThick + parInsulationThickness)) * Cos(parAngle), (parFacetoCenterofTrimmedE - (flangeThick + parInsulationThickness)) * Sin(parAngle), 0
    End If
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
         parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
         parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set (parFacetoCenterofTrimmedE - NEGLIGIBLE_THICKNESS) * Cos(parAngle), (parFacetoCenterofTrimmedE - NEGLIGIBLE_THICKNESS) * Sin(parAngle), 0
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
       
' Set the output
    m_OutputColl.AddOutput "InsNozzle2", ObjInsulatedNozzle2
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set ObjInsulatedNozzle2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
