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
'   Author:         RUK
'   Creation Date:  Wednesday, May 16 2007
'   Description:
'       Planar U Siphon (www.sika.net/pdf/englisch/PRESS3_39_40.PDF)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   22.May.2007     RUK     CR-113569: Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const NEGLIGIBLE_VALUE = 0.0001
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI       As Double

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
    
    Dim parLength As Double
    Dim parOutletHeight As Double
    Dim parInletHeight As Double
    Dim parRadius As Double
    Dim parInsulationThickness As Double
    
'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parLength = arrayOfInputs(2)
    parOutletHeight = arrayOfInputs(3)
    parInletHeight = arrayOfInputs(4)
    parRadius = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim dInsulationDiameter As Double

'Origin is taken at the
'Do initial calculation for preparing the symbol geometry
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, sptOffset2, depth2
    
    If CmpDblLessThan(flangeThick, NEGLIGIBLE_VALUE) Then flangeThick = NEGLIGIBLE_VALUE
    If CmpDblLessThan(flangeThick2, NEGLIGIBLE_VALUE) Then flangeThick2 = NEGLIGIBLE_VALUE

'   Check the insulation thickness whether it exceeds the value of radius of the elbow and radius of the circular bend.
    Dim dInsulationThickness As Double
    dInsulationThickness = parInsulationThickness
    If CmpDblGreaterthanOrEqualTo(dInsulationThickness, 1.5 * pipeDiam / 2) Or _
                CmpDblGreaterthanOrEqualTo(dInsulationThickness, parRadius) Then
        If CmpDblGreaterthanOrEqualTo(parRadius, 1.5 * pipeDiam / 2) Then
            dInsulationThickness = (1.5 * pipeDiam / 2) - 0.001
        Else
            dInsulationThickness = parRadius - 0.001
        End If
    End If
    dInsulationDiameter = pipeDiam + 2 * dInsulationThickness
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Dim oStPoint As New AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    Dim oCenPoint As New AutoMath.DPosition
    
'Create the Cylinder 1 (Output 1)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -(parOutletHeight - pipeDiam - parRadius), 0
'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
    
'Create the 180 degree Bend (Output 2)
    Dim oCircle As IngrGeom3D.Circle3d
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oEnPoint.x, oEnPoint.y, oEnPoint.z, _
                                                             0, 1, 0, dInsulationDiameter / 2)
    'Revolve the circle for 180 degree about the Z axis passing by the origin
    Dim objBend1 As Object
    Dim oRevCen As New AutoMath.DPosition 'revolution center point
    Dim oRevVec As New AutoMath.DVector   'rotation vector for rotation
    oRevCen.Set -(pipeDiam / 2 + parRadius), oEnPoint.y, pipeDiam / 2
    oRevVec.Set 0, 0, -1
    Set objBend1 = PlaceRevolution(m_OutputColl, oCircle, oRevVec, oRevCen, PI, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBend1
    Set objBend1 = Nothing
    
'Create the Cylinder 2 (Output 3)
    oStPoint.Set -(pipeDiam / 2 + 2 * parRadius + pipeDiam / 2), -(parOutletHeight - pipeDiam - parRadius), 0
    oEnPoint.Set oStPoint.x, -(parOutletHeight - parInletHeight + pipeDiam / 2 + 1.5 * pipeDiam / 2), oStPoint.z
'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
    
'Create 90 Degree Bend (Output 4)
    'The Bend diameter is assumed as 1.5*pipeDiam
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oEnPoint.x, oEnPoint.y, oEnPoint.z, _
                                                             0, 1, 0, dInsulationDiameter / 2)
'Revolve the circle for 90 degree about the Z axis
    Dim objElbow As Object
    oRevCen.Set oEnPoint.x - pipeDiam / 2 - 1.5 * pipeDiam / 2, oEnPoint.y, pipeDiam / 2
    oRevVec.Set 0, 0, 1
    Set objElbow = PlaceRevolution(m_OutputColl, oCircle, oRevVec, oRevCen, PI / 2, True)
'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objElbow
    Set objElbow = Nothing
    Set oCircle = Nothing
    Set oRevCen = Nothing
    Set oRevVec = Nothing
    
'Create Cylinder 3 (Output 5)
    oStPoint.Set oEnPoint.x - pipeDiam / 2 - 1.5 * pipeDiam / 2, -(parOutletHeight - parInletHeight), 0
    oEnPoint.Set -parLength, oStPoint.y, 0
'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
    
' Insert your code for output 6(Insulation for Port 1)
    Dim objInsPort1 As Object
    oStPoint.Set -parLength, -(parOutletHeight - parInletHeight), 0
    oEnPoint.Set oStPoint.x + flangeThick + parInsulationThickness, _
                    -(parOutletHeight - parInletHeight), 0
    If CmpDblGreaterthan(flangeDiam1, pipeDiam) Then
        dInsulationDiameter = flangeDiam1 + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
    Set objInsPort1 = Nothing
    
' Insert your code for output 7(Insulation for Port 2)
    Dim objInsPort2 As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set oStPoint.x, -flangeThick2 - parInsulationThickness, 0

    If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then
        dInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set objInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
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
