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

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Const NEGLIGIBLE_VALUE = 0.0001
Private m_oSymGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    Set m_oSymGeomHelper = New SymbolServices
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
    
    'Do initial calculation for preparing the symbol geometry
    Dim dNozzle1Length As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, sptOffset2, depth2
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition

    'Create the Cylinder 1 (Output 1)
    oStPoint.Set 0, 0, 0.5 * (parUDiameter + pipeDiam)
    oEnPoint.Set 0, parOutletHeight - pipeDiam - parUDiameter / 2, 0.5 * (parUDiameter + pipeDiam)
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, pipeDiam
    
    'Create 180 Degree Bend (Output 2)
    Dim oNormal As New AutoMath.DVector
    Dim oMajor As New AutoMath.DVector

    oCenPoint.Set 0, 0, 0
    oNormal.Set 1, 0, 0
    oMajor.Set 0, 0, 1
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateTorus arrayOfOutputs(iOutput), oCenPoint, oNormal, oMajor, _
                    0.5 * (parUDiameter + pipeDiam), 180, pipeDiam
    Set oNormal = Nothing
    Set oMajor = Nothing
    
    'Create the Cylinder 2 (Output 3)
    oStPoint.Set 0, 0, -0.5 * (parUDiameter + pipeDiam)
    oEnPoint.Set 0, parInletHeight - pipeDiam - parUDiameter / 2 - parRadius - pipeDiam / 2, _
                -0.5 * (parUDiameter + pipeDiam)
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, pipeDiam
    
    'Create 90 Degree Bend (Output 4)
    Dim oCircle As IngrGeom3D.Circle3d
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oEnPoint.x, oEnPoint.y, oEnPoint.z, _
                                                            0, 1, 0, pipeDiam / 2)

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
    
    'Create Cylinder 3 (Output 5)
    oStPoint.Set -parInletFacetoOutletCL, parInletHeight - parUDiameter / 2 - pipeDiam, -(parUDiameter / 2 + pipeDiam / 2)
    oEnPoint.Set -(pipeDiam / 2 + parRadius), oStPoint.y, oStPoint.z
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, pipeDiam
        
    'Insert your code for output 6(Clamp muff connection at Port2)
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d

    Dim dAngle As Double
    Dim iCount As Integer
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    dAngle = PI / 3
    oStPoint.Set 0, parOutletHeight - pipeDiam - parUDiameter / 2, 0.5 * (parUDiameter + pipeDiam)

    'First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = oStPoint.x + (flangeDiam2 / 2) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints(3 * iCount - 2) = oStPoint.y - flangeThick2
        LineStrPoints(3 * iCount - 1) = oStPoint.z + (flangeDiam2 / 2) * Sin(PI / 6 + iCount * dAngle)
    Next iCount

    'Assumption: Projection Length is taken as 2 times the flange thickness
    Dim dProjLen As Double
    dProjLen = 2 * flangeThick2
    If CmpDblLessThan(dProjLen, NEGLIGIBLE_VALUE) Then dProjLen = NEGLIGIBLE_VALUE
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    oAxisVect.Set 0, 1, 0
    Dim ObjClampMuffConnection As Object
    Set ObjClampMuffConnection = PlaceProjection(m_OutputColl, oLineString, oAxisVect, dProjLen, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "ClampMuffConnection", ObjClampMuffConnection
    Set ObjClampMuffConnection = Nothing
        
    'Remove references
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oGeomFactory = Nothing
    
    'Place Nozzle 1 (output 7)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -(parInletFacetoOutletCL - sptOffset1 + depth1), (parInletHeight - parUDiameter / 2 - pipeDiam), -(parUDiameter / 2 + pipeDiam / 2)
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzle1Length)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2 (output 8)
    oPlacePoint.Set 0, (parOutletHeight - pipeDiam - parUDiameter / 2 + flangeThick2 + sptOffset2 - depth2), (parUDiameter / 2 + pipeDiam / 2)
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oAxisVect = Nothing
    Set oLineString = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
