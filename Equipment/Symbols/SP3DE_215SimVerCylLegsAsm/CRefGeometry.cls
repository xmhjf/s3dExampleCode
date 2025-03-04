VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CRefGeometry.cls
'   Author:    VRK
'   Creation Date:  FriDay,March 23 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'       ControlPoint at the origin of equipment
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   24.Sep.2007     VRG                 Control point position is adjusted to correct location
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "RefGeometry:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double

    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportReferencePoint  As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parSupportAngularLocation = arrayOfInputs(6)    'P4
    parNumberOfSupports = arrayOfInputs(7)          'P5
    parSupportHeight = arrayOfInputs(8)             'P6
    parSupportLength = arrayOfInputs(9)             'P7
    parSupportThickness = arrayOfInputs(10)         'P8
    parSupportReferencePoint = arrayOfInputs(11)    'P9
    parEndHeadConeTopDiameter = arrayOfInputs(12)   'P51
    parEndHeadConeHeight = arrayOfInputs(13)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(14)     'P51 for TORC and TORS
    parEndHeadDomeradius = arrayOfInputs(15)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(16)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(17)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(18)   'P50
    parInsulationThickness = arrayOfInputs(19)
    
    iOutput = 0
   
    'Insert your code for Control Point of vessel
    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, -parSupportReferencePoint, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "SimVerCylEqpLegControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub



