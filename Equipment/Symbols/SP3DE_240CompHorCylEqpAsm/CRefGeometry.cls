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
'   Author:         VRK
'   Creation Date:  Friday,March 9 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'       ControlPoint at the origin of equipment
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    Dim iOutput     As Double
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parBootHorLocation As Double
    Dim parBootVertLocation As Double
    Dim parBootDiameter As Double
    Dim parBootAngularLocation As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parThirdSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parStiffenerRadius As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
   
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parBootHorLocation = arrayOfInputs(6)           'P4
    parBootVertLocation = arrayOfInputs(7)          'P5
    parBootDiameter = arrayOfInputs(8)              'P6
    parBootAngularLocation = arrayOfInputs(9)       'P7
    parVesselCenterHeight = arrayOfInputs(10)       'P8
    parFirstSupportLocation = arrayOfInputs(11)     'P9
    parSecondSupportLocation = arrayOfInputs(12)    'P10
    parThirdSupportLocation = arrayOfInputs(13)     'P11
    parSupportThickness = arrayOfInputs(14)         'P12
    parSupportLength = arrayOfInputs(15)            'P13
    parStiffenerRadius = arrayOfInputs(16)          'P14
    parEndHeadConeTopDiameter = arrayOfInputs(17)   'P51
    parEndHeadConeHeight = arrayOfInputs(18)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(19)     'P51 for TORC n TORS
    parEndHeadDomeradius = arrayOfInputs(20)        'P50
    parEndHeadFlangedThick1 = arrayOfInputs(21)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(22)     'P51
    parEndHeadSphericalRadius = arrayOfInputs(23)   'P50
    parInsulationThickness = arrayOfInputs(24)
    parXBoltHole = arrayOfInputs(25)
    parYBoltHole = arrayOfInputs(26)
        
    iOutput = 0

    'Insert your code for Control Point of vessel
    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "ComplexHorCylEqpControlPoint", m_oControlPoint
    
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
   
   Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


