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
'   File:           CPhysical.cls
'   Author:         CSA, Inc.
'   Creation Date:  Friday, April 20 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Six Outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private Const UNITSTRANS = 0.0254

Private Const E_FAIL = &H80004005

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    Set m_oGeomHelper = New SymbolServices
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo Errx

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Integer
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl

    Dim parFaceToFace As Double
    Dim parBodyHeight As Double
    Dim parSupportWidth As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFaceToFace = arrayOfInputs(2)
    parBodyHeight = arrayOfInputs(3)
    parSupportWidth = arrayOfInputs(4)
    parXBoltHole = arrayOfInputs(5)
    parYBoltHole = arrayOfInputs(6)

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oPos1 As IJDPosition
    Dim oPos2 As IJDPosition
    Set oPos1 = New DPosition
    Set oPos2 = New DPosition

    'Create the Bottom Small Cylinder (Output 1)
    oPos1.Set 0, 0, 0
    oPos2.Set parBodyHeight, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, 0.65 * parFaceToFace

    'Create the Bottom Cylinder (Output 2)
    oPos1.Set 0.1 * parBodyHeight, 0, 0
    oPos2.Set 0.9 * parBodyHeight, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, parFaceToFace

    Set oPos1 = Nothing
    Set oPos2 = Nothing

    'Create Base Plate (Output 3)
    Dim ObjBasePlateColl As Collection
    Dim oTopSurPts(0 To 3) As IJDPosition
    Dim oBotSurPts(0 To 3) As IJDPosition
    Dim iCount As Integer
    
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = New DPosition
        Set oBotSurPts(iCount) = New DPosition
    Next iCount
    
    oTopSurPts(0).Set 0, 0.5 * parSupportWidth, -0.5 * parSupportWidth
    oTopSurPts(1).Set oTopSurPts(0).x, oTopSurPts(0).y, 0.5 * parSupportWidth
    oTopSurPts(2).Set oTopSurPts(0).x, -0.5 * parSupportWidth, oTopSurPts(1).z
    oTopSurPts(3).Set oTopSurPts(0).x, oTopSurPts(2).y, oTopSurPts(0).z
    
    oBotSurPts(0).Set -2.5 * UNITSTRANS, oTopSurPts(0).y, oTopSurPts(0).z
    oBotSurPts(1).Set -2.5 * UNITSTRANS, oTopSurPts(1).y, oTopSurPts(1).z
    oBotSurPts(2).Set -2.5 * UNITSTRANS, oTopSurPts(2).y, oTopSurPts(2).z
    oBotSurPts(3).Set -2.5 * UNITSTRANS, oTopSurPts(3).y, oTopSurPts(3).z

    Set ObjBasePlateColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjBasePlateColl.Count
        m_OutputColl.AddOutput "BasePlate_", ObjBasePlateColl(iCount)
    Next iCount
    Set ObjBasePlateColl = Nothing
    
    'Create the Edges and Points for Base Plate (Output 4)
    Dim ObjEdgeColl As New Collection
    
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjEdgeColl.Count
        m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
    Next iCount
    Set ObjEdgeColl = Nothing
    
    'Creating the Points on each Face of the Base Plate (Output 5)
    Dim ObjPointColl As New Collection
    
    'Points on Right and Left surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            (oTopSurPts(1).x + oBotSurPts(2).x) / 2, (oTopSurPts(1).y + oBotSurPts(2).y) / 2, (oTopSurPts(1).z + oBotSurPts(2).z) / 2)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            (oTopSurPts(0).x + oBotSurPts(3).x) / 2, (oTopSurPts(0).y + oBotSurPts(3).y) / 2, (oTopSurPts(0).z + oBotSurPts(3).z) / 2)
    'Points on Front and Back surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            (oTopSurPts(3).x + oBotSurPts(2).x) / 2, (oTopSurPts(3).y + oBotSurPts(2).y) / 2, (oTopSurPts(3).z + oBotSurPts(2).z) / 2)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            (oTopSurPts(0).x + oBotSurPts(1).x) / 2, (oTopSurPts(0).y + oBotSurPts(1).y) / 2, (oTopSurPts(0).z + oBotSurPts(1).z) / 2)
    'Points on Top and Bottom surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            (oTopSurPts(0).x + oTopSurPts(2).x) / 2, (oTopSurPts(0).y + oTopSurPts(2).y) / 2, (oTopSurPts(0).z + oTopSurPts(2).z) / 2)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            (oBotSurPts(0).x + oBotSurPts(2).x) / 2, (oBotSurPts(0).y + oBotSurPts(2).y) / 2, (oBotSurPts(0).z + oBotSurPts(2).z) / 2)
    'Set the output
    iOutput = iOutput + 1
    For iCount = 1 To ObjPointColl.Count
        m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
    Next iCount
    Set ObjPointColl = Nothing
    
    'Remove the References
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = Nothing
        Set oBotSurPts(iCount) = Nothing
    Next iCount
    Set oGeomFactory = Nothing
    
    'Create Equipment Foundation Port at the Bottom of Base Plate (Output 6)
    '
    '                      |-------|
    '   Z                  |       |
    '   ^                  |       |
    '   |                  |   X   |
    '   |                  |   ^   |
    '   |                  |   |   |
    '   -----> Y           |   |   |Port CS
    '   Symbol CS          |   |----> Y
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |-------|
    
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = -2.5 * UNITSTRANS
    dOrigin(1) = 0
    dOrigin(2) = 0
    dXaxis(0) = 0
    dXaxis(1) = 0
    dXaxis(2) = 1

    dZaxis(0) = -1
    dZaxis(1) = 0#
    dZaxis(2) = 0

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
   
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub
