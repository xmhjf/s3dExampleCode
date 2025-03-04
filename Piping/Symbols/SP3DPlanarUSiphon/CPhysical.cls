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
'   Author:         RUK
'   Creation Date:  Wednesday, May 16 2007
'   Description:
'       Planar U Siphon (www.sika.net/pdf/englisch/PRESS3_39_40.PDF)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   22.May.2007     RUK     CR-113569: Created the symbol.
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
    
    Dim parLength As Double
    Dim parOutletHeight As Double
    Dim parInletHeight As Double
    Dim parRadius As Double
'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parLength = arrayOfInputs(2)
    parOutletHeight = arrayOfInputs(3)
    parInletHeight = arrayOfInputs(4)
    parRadius = arrayOfInputs(5)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
'Origin is taken at the Outlet Port
'Do initial calculation for preparing the symbol geometry
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, sptOffset2, depth2
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    Dim oCenPoint As New AutoMath.DPosition
    
'Create the Cylinder 1 (Output 1)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -(parOutletHeight - pipeDiam - parRadius), 0
'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, pipeDiam
    
'Create 180 Degree Bend (Output 2)
    Dim oNormal As New AutoMath.DVector
    Dim oMajor As New AutoMath.DVector
    oCenPoint.Set -(parRadius + pipeDiam / 2), oEnPoint.y, 0
    oNormal.Set 0, 0, -1
    oMajor.Set 1, 0, 0
'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateTorus arrayOfOutputs(iOutput), oCenPoint, oNormal, oMajor, parRadius + pipeDiam / 2, _
                                                        180, pipeDiam
    Set oNormal = Nothing
    Set oMajor = Nothing
    
'Create the Cylinder 2 (Output 3)
    oStPoint.Set -(pipeDiam / 2 + 2 * parRadius + pipeDiam / 2), -(parOutletHeight - pipeDiam - parRadius), 0
    oEnPoint.Set oStPoint.x, -(parOutletHeight - parInletHeight + pipeDiam / 2 + 1.5 * pipeDiam / 2), oStPoint.z
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, pipeDiam
    
'Create 90 Degree Bend (Output 4)
    'The Bend diameter is assumed as 1.5*pipeDiam
    Dim oCircle As IngrGeom3D.Circle3d
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oEnPoint.x, oEnPoint.y, oEnPoint.z, _
                                                            0, 1, 0, pipeDiam / 2)
    'Revolve the circle for 90 degree about the Z axis
    Dim objElbow As Object
    Dim oRevCen As New AutoMath.DPosition 'revolution center point
    Dim oRevVec As New AutoMath.DVector   'revolution vector for rotation
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
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, pipeDiam

'Remove references
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oGeomFactory = Nothing
    
'Place Inlet Nozzle (output 6)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parLength - sptOffset1 + depth1, -(parOutletHeight - parInletHeight), 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'Place Outlet Nozzle (output 7)
    oPlacePoint.Set 0, sptOffset2 - depth2, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
