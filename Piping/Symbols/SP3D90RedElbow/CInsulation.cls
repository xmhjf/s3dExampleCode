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
'   Author:         MS
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: F84 Page No. D-56 of PDS Piping Component Data Reference Guide.
'    Symbol is created with Seven Outputs
'    The Three Insulation aspect outputs are created as follows:
'    ObjInsulatedBody using PlaceRevolution, ObjReducingBodyIns using PlaceCone and
'    ObjTangentBodyIns using PlaceCylinder
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006     KKC                     DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.Aug.2007     VRG                     TR-124960 Insulation diameter was computed based on End preparation
'  07.Nov.2007      MA                      TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                                       in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages
Private Sub Class_Initialize()

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
    Dim cptOffset2       As Double
    Dim depth2           As Double
    
    Dim ObjInsulatedNozzle1 As Object
    Dim ObjInsulatedNozzle2 As Object
    Dim iOutput     As Integer
    Dim ObjInsulatedBody As Object
    Dim ObjReducingBodyIns As Object
    Dim ObjTangentBodyIns As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parReducingLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parReducingLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    iOutput = 0

 ' Insert your code for output 3(Insulated Body)
    Dim ElbowRadius As Double
    Dim InsulationRadius As Double
    Dim PI           As Double
    PI = 4 * Atn(1)
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset2, depth2

    ElbowRadius = parFace1toCenter - parReducingLength
    
' Construct a circle that will be used to create the revolution
    InsulationRadius = pipeDiam2 / 2 + parInsulationThickness
'   Ensure that the InsulationRadius does not exceed the InsulationRadius
    If CmpDblGreaterthan(InsulationRadius, ElbowRadius) Then InsulationRadius = ElbowRadius
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                -parFace1toCenter + parReducingLength, 0, 0, _
                                                1, 0, 0, _
                                                InsulationRadius)
' Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ElbowRadius, ElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircle, CV, CP, PI / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set geomFactory = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing

 'Insert your code for output 4(Reducing Body Insulation)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + parReducingLength, 0, 0
    Set ObjReducingBodyIns = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2 + parInsulationThickness, _
                                                                    pipeDiam2 / 2 + parInsulationThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReducingBodyIns
    Set ObjReducingBodyIns = Nothing

'Insert your code for output 5(Tangent Body Insulation)
    Dim dInsulationDia As Double
    stPoint.Set 0, parFace1toCenter - parReducingLength, 0
    enPoint.Set 0, parFace2toCenter, 0
    dInsulationDia = pipeDiam2 + 2 * parInsulationThickness

    Set ObjTangentBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangentBodyIns
    Set ObjTangentBodyIns = Nothing
           
' Insert your code for output 8(Insulated Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    stPoint.Set -parFace1toCenter, 0, 0
    If CmpDblGreaterthanOrEqualTo(-parFace1toCenter + flangeThick + parInsulationThickness, 0) Then
        enPoint.Set 0, 0, 0
    Else
        enPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    End If
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam + 2 * parInsulationThickness
    End If
   
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set -parFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
    End If
    
    Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    m_OutputColl.AddOutput "InsulatedNozzle1", ObjInsulatedNozzle1
    Set ObjInsulatedNozzle1 = Nothing

' Insert your code for output 9(Insulated Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset2, depth2
    stPoint.Set 0, parFace2toCenter, 0
    If CmpDblLessThanOrEqualTo(parFace2toCenter - (flangeThick + parInsulationThickness), 0) Then
        enPoint.Set 0, 0, 0
    Else
        enPoint.Set 0, parFace2toCenter - (flangeThick + parInsulationThickness), 0
    End If

    If CmpDblGreaterthan(flangeDiam, pipeDiam2) Then
         dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Else
         dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
    End If
    
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set 0, parFace2toCenter - NEGLIGIBLE_THICKNESS, 0
    End If
  
    Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    m_OutputColl.AddOutput "InsulatedNozzle2", ObjInsulatedNozzle2
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set ObjInsulatedNozzle2 = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
