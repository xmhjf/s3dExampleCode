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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:           ACM
'   Creation Date:  Thursday, Nov 14 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols.
'  18.Jul.2007     dkl  TR-124260 Rectified insulation of the ports to enable placement when flange thickness is zero for value of flange dia.
'                                 Used appropriate CmpdblXXX functions where required.
'   7.Nov.2007      VRK  TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                               in which case small cylinders of negligible thickness will be created
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
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object

    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim InsulationRadius As Double
    Dim ElbowRadius As Double
    Dim Angle As Double
    Angle = PI / 6

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    'If flange thickness is zero, it is assigned a negligible value to prevent port insulation from failing.
    
    ElbowRadius = parFacetoCenter / Tan(Angle / 2)
    InsulationRadius = pipeDiam / 2 + parInsulationThickness
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -parFacetoCenter, ElbowRadius, 0
    CV.Set 0, 0, 1
' Limit Insulation until InsulationRadius coincides with the Radius of the Elbow
    If CmpDblLessThanOrEqualTo(InsulationRadius, ElbowRadius) Then
'     Construct a circle that will be used to create the revolution
        Dim objCircleI   As IngrGeom3D.Circle3d
        Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             -parFacetoCenter, 0, 0, _
                            1, 0, 0, _
                          InsulationRadius)
           ' Revolve the circle for 90 degree about the Z axis passing by the origin
        
        
        Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, Angle, True)
                    Set objCircleI = Nothing
    Else
        Dim HalfCord As Double
        HalfCord = Sqr(InsulationRadius ^ 2 - ElbowRadius ^ 2)

        Dim points(1 To 13)   As New AutoMath.DPosition
        Dim objBsplineI  As IngrGeom3D.BSplineCurve3d
        points(1).Set -parFacetoCenter, ElbowRadius, 0
        points(2).Set -parFacetoCenter, ElbowRadius, HalfCord
        points(3).Set -parFacetoCenter, ElbowRadius / 2, InsulationRadius
        points(4).Set -parFacetoCenter, 0, InsulationRadius
        points(5).Set -parFacetoCenter, -InsulationRadius / 2, InsulationRadius
        points(6).Set -parFacetoCenter, -InsulationRadius, InsulationRadius / 2
        points(7).Set -parFacetoCenter, -InsulationRadius, 0
        points(8).Set -parFacetoCenter, -InsulationRadius, -InsulationRadius / 2
        points(9).Set -parFacetoCenter, -InsulationRadius / 2, -InsulationRadius
        points(10).Set -parFacetoCenter, 0, -InsulationRadius
        points(11).Set -parFacetoCenter, ElbowRadius / 2, -InsulationRadius
        points(12).Set -parFacetoCenter, ElbowRadius, -HalfCord
        points(13).Set -parFacetoCenter, ElbowRadius, 0


        Set objBsplineI = PlaceTrBspline(5, points)
        ' Revolve the Bspline for 90 degree about the Z axis passing by the origin
    
    
        Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objBsplineI, _
                    CV, CP, Angle, True)
                    Set objBsplineI = Nothing
    End If

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
     
      ' Insert your code for output 2(Insulated Port1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    stPoint.Set -parFacetoCenter, 0, 0
    
    If CmpDblLessThanOrEqualTo(flangeDiam, 0) Then
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        enPoint.Set -parFacetoCenter * 0.99, 0, 0
    Else
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        enPoint.Set -(parFacetoCenter - flangeThick - parInsulationThickness), 0, 0
    End If
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set -(parFacetoCenter - NEGLIGIBLE_THICKNESS), 0, 0
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

 ' Insert your code for output 3(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    'If flange thickness is zero, it is assigned a negligible value to prevent port insulation from failing.
    
    stPoint.Set parFacetoCenter * Cos(Angle), parFacetoCenter * Sin(Angle), 0
    If CmpDblLessThanOrEqualTo(flangeDiam, 0) Then
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        enPoint.Set parFacetoCenter * 0.99 * Cos(Angle), parFacetoCenter * 0.99 * Sin(Angle), 0
    Else
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        enPoint.Set (parFacetoCenter - flangeThick - parInsulationThickness) * Cos(Angle), (parFacetoCenter - flangeThick - parInsulationThickness) * Sin(Angle), 0
    End If
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set (parFacetoCenter - NEGLIGIBLE_THICKNESS) * Cos(Angle), (parFacetoCenter - NEGLIGIBLE_THICKNESS) * Sin(Angle), 0
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
