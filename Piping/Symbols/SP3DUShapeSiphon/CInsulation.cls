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
'   Author:         VRK
'   Creation Date:  Friday, May 18 2007
'   Description:
'       U Shape Siphon (Source: http://www.sika.net/pdf/englisch/PRESS3_39_40.PDF)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   23.May.2007     VRK     CR-113569: Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const NEGLIGIBLE_VALUE = 0.0001
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI  As Double

Private Sub Class_Initialize()
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam         As Double
    Dim flangeThick      As Double
    Dim flangeDiam1       As Double
    Dim flangeDiam2       As Double
    Dim sptOffset1       As Double
    Dim depth1            As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    
    Dim parInletFacetoOutletCL As Double
    Dim parOutletHeight As Double
    Dim parInletHeight As Double
    Dim parRadius As Double
    Dim parUDiameter As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInletFacetoOutletCL = arrayOfInputs(2)
    parOutletHeight = arrayOfInputs(3)
    parInletHeight = arrayOfInputs(4)
    parRadius = arrayOfInputs(5)
    parUDiameter = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim dInsulationDiameter As Double

    'Do initial calculation for preparing the symbol geometry
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, sptOffset2, depth2
    
    If CmpDblLessThan(flangeThick, NEGLIGIBLE_VALUE) Then flangeThick = NEGLIGIBLE_VALUE
    If CmpDblLessThan(flangeThick2, NEGLIGIBLE_VALUE) Then flangeThick2 = NEGLIGIBLE_VALUE

    'Check the insulation thickness whether it exceeds the value of radius of the elbow and radius of the circular bend.
    Dim dInsulationThickness As Double
    dInsulationThickness = parInsulationThickness
    
    If CmpDblGreaterthanOrEqualTo(parInsulationThickness, parUDiameter / 2) Or _
                            CmpDblGreaterthanOrEqualTo(parInsulationThickness, parRadius) Then
        If CmpDblGreaterthanOrEqualTo(parUDiameter / 2, parRadius) Then
            dInsulationThickness = parRadius - 0.001
        Else
            dInsulationThickness = (parUDiameter / 2) - 0.001
        End If
    End If
    
    dInsulationDiameter = pipeDiam + 2 * dInsulationThickness
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
  
    'Create the Cylinder 1 (Output 9)
    oStPoint.Set 0, 0, 0.5 * (parUDiameter + pipeDiam)
    oEnPoint.Set 0, parOutletHeight - pipeDiam - parUDiameter / 2, 0.5 * (parUDiameter + pipeDiam)
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
    
    'Create 180 Degree Bend (Output 10)
    Dim oNormal As New AutoMath.DVector
    Dim oMajor As New AutoMath.DVector

    oCenPoint.Set 0, 0, 0
    oNormal.Set 1, 0, 0
    oMajor.Set 0, 0, 1
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateTorus arrayOfOutputs(iOutput), oCenPoint, oNormal, oMajor, 0.5 * (parUDiameter + pipeDiam), _
                                                        180, dInsulationDiameter
    Set oNormal = Nothing
    Set oMajor = Nothing
    
    'Create the Cylinder 2 (Output 11)
    oStPoint.Set 0, 0, -0.5 * (parUDiameter + pipeDiam)
    oEnPoint.Set 0, parInletHeight - pipeDiam - parUDiameter / 2 - parRadius - pipeDiam / 2, -0.5 * (parUDiameter + pipeDiam)
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
    
    'Create 90 Degree Bend (Output 12)
    Dim oCircle As IngrGeom3D.Circle3d
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oEnPoint.x, oEnPoint.y, oEnPoint.z, _
                                                            0, 1, 0, dInsulationDiameter / 2)

    'Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim objElbow As Object
    Dim oRevCen As AutoMath.DPosition  'revolution center point
    Dim oRevVec As AutoMath.DVector    'rotation vector for rotation
    Set oRevCen = New AutoMath.DPosition
    Set oRevVec = New AutoMath.DVector
    oRevCen.Set -(pipeDiam / 2 + parRadius), parInletHeight - pipeDiam - parUDiameter / 2 - parRadius - pipeDiam / 2, -(parUDiameter / 2 + pipeDiam / 2)
    oRevVec.Set 0, 0, 1
    Set objElbow = PlaceRevolution(m_OutputColl, oCircle, oRevVec, oRevCen, PI / 2, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objElbow
    Set objElbow = Nothing
    Set oCircle = Nothing
    Set oRevCen = Nothing
    Set oRevVec = Nothing

    'Create Cylinder 3 (Output 13)
    oStPoint.Set -parInletFacetoOutletCL, parInletHeight - parUDiameter / 2 - pipeDiam, -(parUDiameter / 2 + pipeDiam / 2)
    oEnPoint.Set -(pipeDiam / 2 + parRadius), oStPoint.y, oStPoint.z
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
    
    'Insert your code for output 14(Clamp muff connection at Port2)
    Dim ObjInsClampMuffConnection As Object
    'Assumptions:
    ' 1. The insulation shape for the Clamp muff connection is considered to be cylindrical
    ' 2. Projection Length is taken as 2 times the flange thickness
    Dim dProjLen As Double
    dProjLen = 2 * flangeThick2
    If CmpDblLessThan(dProjLen, NEGLIGIBLE_VALUE) Then dProjLen = NEGLIGIBLE_VALUE
    oStPoint.Set 0, parOutletHeight - pipeDiam - parUDiameter / 2 + flangeThick2, 0.5 * (parUDiameter + pipeDiam)
    oEnPoint.Set oStPoint.x, oStPoint.y - dProjLen, oStPoint.z
    If CmpDblGreaterthanOrEqualTo(flangeDiam2, pipeDiam) Then
        dInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
    Else
         dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set ObjInsClampMuffConnection = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
             dInsulationDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "InsClampMuffConnection", ObjInsClampMuffConnection
    Set ObjInsClampMuffConnection = Nothing

    'Insert your code for output 15(Insulation for Port 1)
    Dim objInsPort1 As Object
    oStPoint.Set -(parInletFacetoOutletCL), (parInletHeight - parUDiameter / 2 - pipeDiam), -(parUDiameter / 2 + pipeDiam / 2)
    oEnPoint.Set oStPoint.x + flangeThick + parInsulationThickness, _
                    (parInletHeight - parUDiameter / 2 - pipeDiam), -(parUDiameter / 2 + pipeDiam / 2)

    If CmpDblGreaterthan(flangeDiam1, pipeDiam) Then
        dInsulationDiameter = flangeDiam1 + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

    Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
    Set objInsPort1 = Nothing
    
    'Insert your code for output 16(Insulation for Port 2)
    Dim objInsPort2 As Object
    oStPoint.Set 0, parOutletHeight - pipeDiam - parUDiameter / 2 + flangeThick2, 0.5 * (parUDiameter + pipeDiam)
    oEnPoint.Set oStPoint.x, oStPoint.y - flangeThick2 - parInsulationThickness, oStPoint.z
    If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then
        dInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set objInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort2
    Set objInsPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Class_Terminate()
   Set m_oSymGeomHelper = Nothing
End Sub
