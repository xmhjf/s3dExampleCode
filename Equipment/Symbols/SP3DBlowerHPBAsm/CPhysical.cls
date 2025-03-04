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
'   Creation Date:  Monday, April 23 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Eight Outputs.
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

    Dim parBottomtoCenter As Double
    Dim parCentertoEndofOutletFlange As Double
    Dim parCentertoOutletCenter As Double
    Dim parFacetoFace As Double
    Dim parCentertoEndofInlet As Double
    Dim parSmallShaftLength As Double
    Dim parInsideInletDiameter As Double
    Dim parHousingWidth As Double
    Dim parDistanceBaseCentertoHole As Double
    Dim parDistanceBaseCentertoEnd As Double
    Dim parShaftDiameter As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBottomtoCenter = arrayOfInputs(2)              'A
    parCentertoEndofOutletFlange = arrayOfInputs(3)   'B
    parCentertoOutletCenter = arrayOfInputs(4)        'D
    parFacetoFace = arrayOfInputs(5)                  'H
    parCentertoEndofInlet = arrayOfInputs(6)          'J
    parSmallShaftLength = arrayOfInputs(7)            'K
    parInsideInletDiameter = arrayOfInputs(8)         'L
    parHousingWidth = arrayOfInputs(9)                'M
    parDistanceBaseCentertoHole = arrayOfInputs(10)   'T
    parDistanceBaseCentertoEnd = arrayOfInputs(11)    'U
    parShaftDiameter = arrayOfInputs(12)              'SD
    parXBoltHole = arrayOfInputs(13)                  'HX
    parYBoltHole = arrayOfInputs(14)                  'HY

    Dim oPos1 As New AutoMath.DPosition
    Dim oPos2 As New AutoMath.DPosition

    'Create the Body Cylinder (Output 1)
    oPos1.Set parCentertoEndofInlet - 0.5 * parHousingWidth, 0, 0
    oPos2.Set parCentertoEndofInlet + 0.5 * parHousingWidth, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, 2 * parCentertoOutletCenter + parHousingWidth

    'Create the Outlet Box (Output 2)
    Dim oBox As Object
    oPos1.Set parCentertoEndofInlet - 0.5 * parHousingWidth, 0.75 * parCentertoEndofOutletFlange, parCentertoOutletCenter - 0.5 * parHousingWidth
    oPos2.Set parCentertoEndofInlet + 0.5 * parHousingWidth, 0, parCentertoOutletCenter + 0.5 * parHousingWidth
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox

    'Create the Inlet Cylinder (Output 3)
    oPos1.Set 0, 0, 0
    oPos2.Set parCentertoEndofInlet - 0.5 * parHousingWidth, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, parInsideInletDiameter

    'Create the Shaft Cylinder (Output 4)
    oPos1.Set parCentertoEndofInlet + 0.5 * parHousingWidth, 0, 0
    oPos2.Set parFacetoFace + parSmallShaftLength, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, parShaftDiameter

    'Create the Nozzle Cone (Output 5)
    Dim oCone As Object
    oPos1.Set parCentertoEndofInlet, parCentertoEndofOutletFlange, parCentertoOutletCenter
    oPos2.Set parCentertoEndofInlet, 0.75 * parCentertoEndofOutletFlange, parCentertoOutletCenter
    Set oCone = PlaceCone(m_OutputColl, oPos1, oPos2, parHousingWidth / 2, 0.7 * parHousingWidth / 2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCone
    Set oCone = Nothing

    'Create the Bottom Base (Output 6)
    oPos1.Set parFacetoFace, -parDistanceBaseCentertoHole, 0.5 * UNITSTRANS - parBottomtoCenter
    oPos2.Set parCentertoEndofInlet + 0.5 * parHousingWidth, parDistanceBaseCentertoHole, -0.2 * parBottomtoCenter
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox

    'Create the Base Plate (Output 7)
    oPos1.Set parCentertoEndofInlet + 0.5 * parHousingWidth, -parDistanceBaseCentertoEnd, -parBottomtoCenter
    oPos2.Set parFacetoFace, parDistanceBaseCentertoEnd, -parBottomtoCenter + 0.5 * UNITSTRANS
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox

    Set oBox = Nothing
    Set oPos1 = Nothing
    Set oPos2 = Nothing
    
    'Create Equipment Foundation Port at the Bottom of Base Plate (Output 8)
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
    dOrigin(0) = parFacetoFace - (parFacetoFace - (parCentertoEndofInlet + 0.5 * parHousingWidth)) / 2
    dOrigin(1) = 0
    dOrigin(2) = -parBottomtoCenter
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
