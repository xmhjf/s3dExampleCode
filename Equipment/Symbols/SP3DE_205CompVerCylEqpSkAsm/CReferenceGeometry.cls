VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:03 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         RUK
'   Creation Date:  Monday, March 26 2007
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
Private Const MODULE = "RefGeom:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim parVesselHeight1 As Double
    Dim parVesselDiameter1 As Double
    Dim parSupportHeight1 As Double
    Dim parVesselHeight2 As Double
    Dim parVesselDiameter2 As Double
    Dim parSupportHeight2 As Double
    Dim parVesselHeight3 As Double
    Dim parVesselDiameter3 As Double
    Dim parSupportHeight3 As Double
    Dim parVesselHeight4 As Double
    Dim parVesselDiameter4 As Double
    Dim parSkirtHeight As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parVesselReferencePoint As Double
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parEnd3Type As Long
    Dim parEndHead1ConeTopDiameter As Double
    Dim parEndHead1ConeHeight As Double
    Dim parEndHead1KnuckleRadius As Double
    Dim parEndHead1Domeradius As Double
    Dim parEndHead1FlangedThick1 As Double
    Dim parEndHead1FlangedThick2 As Double
    Dim parEndHead1SphericalRadius As Double
    Dim parEndHead2ConeTopDiameter As Double
    Dim parEndHead2ConeHeight As Double
    Dim parEndHead2KnuckleRadius As Double
    Dim parEndHead2Domeradius As Double
    Dim parEndHead2FlangedThick1 As Double
    Dim parEndHead2FlangedThick2 As Double
    Dim parEndHead2SphericalRadius As Double
    Dim parEndHead3ConeTopDiameter As Double
    Dim parEndHead3ConeHeight As Double
    Dim parEndHead3KnuckleRadius As Double
    Dim parEndHead3Domeradius As Double
    Dim parEndHead3FlangedThick1 As Double
    Dim parEndHead3FlangedThick2 As Double
    Dim parEndHead3SphericalRadius As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight1 = arrayOfInputs(2)             'P2
    parVesselDiameter1 = arrayOfInputs(3)           'P3
    parSupportHeight1 = arrayOfInputs(4)            'P4
    parVesselHeight2 = arrayOfInputs(5)             'P5
    parVesselDiameter2 = arrayOfInputs(6)           'P6
    parSupportHeight2 = arrayOfInputs(7)            'P7
    parVesselHeight3 = arrayOfInputs(8)             'P8
    parVesselDiameter3 = arrayOfInputs(9)           'P9
    parSupportHeight3 = arrayOfInputs(10)           'P10
    parVesselHeight4 = arrayOfInputs(11)            'P11
    parVesselDiameter4 = arrayOfInputs(12)          'P12
    parSkirtHeight = arrayOfInputs(13)              'P13
    parSkirtTopDiameter = arrayOfInputs(14)         'P14
    parSkirtBottomDiameter = arrayOfInputs(15)      'P15
    parVesselReferencePoint = arrayOfInputs(16)     'P16
    parEnd1Type = arrayOfInputs(17)
    parEnd2Type = arrayOfInputs(18)
    parEnd3Type = arrayOfInputs(19)
    parEndHead1ConeTopDiameter = arrayOfInputs(20)  'P51
    parEndHead1ConeHeight = arrayOfInputs(21)       'P50
    parEndHead1KnuckleRadius = arrayOfInputs(22)    'P51 for TORC & TORS
    parEndHead1Domeradius = arrayOfInputs(23)       'P50 for Dome
    parEndHead1FlangedThick1 = arrayOfInputs(24)    'P50 for FLGD
    parEndHead1FlangedThick2 = arrayOfInputs(25)    'P51 for FLGD
    parEndHead1SphericalRadius = arrayOfInputs(26)  'P50 for TORS
    parEndHead2ConeTopDiameter = arrayOfInputs(27)  'P51
    parEndHead2ConeHeight = arrayOfInputs(28)       'P50
    parEndHead2KnuckleRadius = arrayOfInputs(29)    'P51 for TORC & TORS
    parEndHead2Domeradius = arrayOfInputs(30)       'P50 for Dome
    parEndHead2FlangedThick1 = arrayOfInputs(31)    'P50 for FLGD
    parEndHead2FlangedThick2 = arrayOfInputs(32)    'P51 for FLGD
    parEndHead2SphericalRadius = arrayOfInputs(33)  'P50 for TORS
    parEndHead3ConeTopDiameter = arrayOfInputs(34)  'P51
    parEndHead3ConeHeight = arrayOfInputs(35)       'P50
    parEndHead3KnuckleRadius = arrayOfInputs(36)    'P51 for TORC & TORS
    parEndHead3Domeradius = arrayOfInputs(37)       'P50 for Dome
    parEndHead3FlangedThick1 = arrayOfInputs(38)    'P50 for FLGD
    parEndHead3FlangedThick2 = arrayOfInputs(39)    'P51 for FLGD
    parEndHead3SphericalRadius = arrayOfInputs(40)  'P50 for TORS
    parInsulationThickness = arrayOfInputs(41)

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    iOutput = 0
    'Origin is at DP/PP 2 at the Top of Skirt.
    'Insert your code for output (Point at origin)
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                  0, 0, -parSkirtHeight, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
