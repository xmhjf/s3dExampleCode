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
'   Creation Date:  Thursday, April 26 2007
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

    Dim parOutsideHeightofHousing As Double
    Dim parOutsideWidthofHousing As Double
    Dim parInletDiameter As Double
    Dim parCentertoBottom As Double
    Dim parDischargeFlangeWidth As Double
    Dim parBottomBoxWidth As Double
    Dim parLength As Double
    Dim parCentertoEndofOutletNozzle As Double
    Dim parCentertoTopofOutlet As Double
    Dim parCentertoBackSideofHousing As Double
    Dim parCentertoEndofHousing As Double
    Dim parCentertoTopofOutletNozzle As Double
    Dim parDistInlettoCenterofHousing As Double
    Dim parFacetoCenterofHousing As Double
    Dim parCentofFantoCentofDisch As Double
    Dim parShaftDiameter As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOutsideHeightofHousing = arrayOfInputs(2)       'A
    parOutsideWidthofHousing = arrayOfInputs(3)        'B
    parInletDiameter = arrayOfInputs(4)                'C
    parCentertoBottom = arrayOfInputs(5)               'D
    parDischargeFlangeWidth = arrayOfInputs(6)         'DX
    parBottomBoxWidth = arrayOfInputs(7)               'GC
    parLength = arrayOfInputs(8)                       'H
    parCentertoEndofOutletNozzle = arrayOfInputs(9)    'HA
    parCentertoTopofOutlet = arrayOfInputs(10)         'HC
    parCentertoBackSideofHousing = arrayOfInputs(11)   'HE
    parCentertoEndofHousing = arrayOfInputs(12)        'HJ
    parCentertoTopofOutletNozzle = arrayOfInputs(13)   'HN
    parDistInlettoCenterofHousing = arrayOfInputs(14)  'J
    parFacetoCenterofHousing = arrayOfInputs(15)       'P
    parCentofFantoCentofDisch = arrayOfInputs(16)      'Q
    parShaftDiameter = arrayOfInputs(17)               'SD
    parXBoltHole = arrayOfInputs(18)                   'HX
    parYBoltHole = arrayOfInputs(19)                   'HY

    Dim oPos1 As New AutoMath.DPosition
    Dim oPos2 As New AutoMath.DPosition
    
    'Create the Body Cylinder (Output 1)
    oPos1.Set parDistInlettoCenterofHousing - 0.5 * parOutsideWidthofHousing, -0.15 * parCentertoBackSideofHousing, 0
    oPos2.Set 0.5 * parOutsideWidthofHousing + parDistInlettoCenterofHousing, -0.15 * parCentertoBackSideofHousing, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, 2 * parCentertoBackSideofHousing

    'Create the Body Box (Output 2)
    Dim oBoxColl As Collection
    Dim oTopSurPts(0 To 3) As IJDPosition
    Dim oBotSurPts(0 To 3) As IJDPosition
    Dim iCount As Integer
    
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = New DPosition
        Set oBotSurPts(iCount) = New DPosition
    Next iCount
    
    oTopSurPts(0).Set parDistInlettoCenterofHousing - 0.5 * parOutsideWidthofHousing, 0.5 * (parCentertoEndofOutletNozzle - parCentertoEndofHousing) + parCentertoEndofHousing, parCentertoTopofOutletNozzle
    oTopSurPts(1).Set 0.5 * parOutsideWidthofHousing + parDistInlettoCenterofHousing, oTopSurPts(0).y, oTopSurPts(0).z
    oTopSurPts(2).Set oTopSurPts(1).x, -0.15 * parCentertoBackSideofHousing, parCentertoBackSideofHousing
    oTopSurPts(3).Set oTopSurPts(0).x, oTopSurPts(2).y, oTopSurPts(2).z
    
    oBotSurPts(0).Set oTopSurPts(0).x, oTopSurPts(0).y, parCentertoTopofOutletNozzle - parOutsideHeightofHousing
    oBotSurPts(1).Set oTopSurPts(1).x, oTopSurPts(1).y, oBotSurPts(0).z
    oBotSurPts(2).Set oTopSurPts(2).x, oTopSurPts(2).y, oBotSurPts(0).z
    oBotSurPts(3).Set oTopSurPts(3).x, oTopSurPts(3).y, oBotSurPts(0).z
    
    Set oBoxColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To oBoxColl.Count
        m_OutputColl.AddOutput "BodyBox_", oBoxColl(iCount)
    Next iCount
    Set oBoxColl = Nothing
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = Nothing
        Set oBotSurPts(iCount) = Nothing
    Next iCount

    'Create the Shaft Cylinder (Output 3)
    oPos1.Set parDistInlettoCenterofHousing + parFacetoCenterofHousing, 0, 0
    oPos2.Set 0.5 * parOutsideWidthofHousing + parDistInlettoCenterofHousing, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, parShaftDiameter
    
    'Create the Small Bottom Box (Output 4)
    Dim oBox As Object
    oPos1.Set parDistInlettoCenterofHousing - 0.5 * parOutsideWidthofHousing, parBottomBoxWidth, -parCentertoBottom
    oPos2.Set 0.5 * (parDistInlettoCenterofHousing - 0.5 * parOutsideWidthofHousing), -parBottomBoxWidth, parInletDiameter - parCentertoBottom
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox

    'Create the Base Plate (Output 5)
    oPos1.Set 0.5 * parOutsideWidthofHousing + parDistInlettoCenterofHousing, parBottomBoxWidth, -parCentertoBottom
    oPos2.Set 0.5 * parLength + 0.5 * parOutsideWidthofHousing + parDistInlettoCenterofHousing, -parBottomBoxWidth, -parShaftDiameter
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    
    Set oBox = Nothing
    Set oPos1 = Nothing
    Set oPos2 = Nothing
    
    'Create Equipment Foundation Port at the Bottom of Base Plate (Output 6)
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |   X   |
    '   |                  |   ^   |
    '   |                  |   |   |
    '   -----> X           |   |   |Port CS
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
    dOrigin(0) = parLength / 2
    dOrigin(1) = 0
    dOrigin(2) = -parCentertoBottom
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

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
