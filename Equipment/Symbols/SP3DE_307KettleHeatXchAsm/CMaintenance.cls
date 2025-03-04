VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         VRK
'   Creation Date:  Monday,March 5 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           BundlePulling Cylider at the Front End
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim ObjBundlePullCylin As Object
    
    Dim parFrontEndType As Long
    Dim parVesselHeadType As Long
    Dim parExchangerLength As Double
    Dim parExchangerNeckLength As Double
    Dim parExchangerNeckDiameter As Double
    Dim parExchangerTaperLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parBundlePullingLength As Double
    Dim parBotSupportCenFromPP As Double
    Dim parBottomSupportCentoCen As Double
    Dim parSupport1Thickness As Double
    Dim parSupport2Thickness As Double
    Dim parBottomSupportHeight As Double
    Dim parSupportLength As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parFrontEndDiameter As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    Dim parFP1XBoltHole As Double
    Dim parFP1YBoltHole As Double
    Dim parFP2XBoltHole As Double
    Dim parFP2YBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndType = arrayOfInputs(2)
    parVesselHeadType = arrayOfInputs(3)
    parExchangerLength = arrayOfInputs(4)           'P1
    parExchangerNeckLength = arrayOfInputs(5)       'P2
    parExchangerNeckDiameter = arrayOfInputs(6)     'P3
    parExchangerTaperLength = arrayOfInputs(7)      'P4
    parExchangerDiameter = arrayOfInputs(8)         'P5
    parExchangerFlangeTk1 = arrayOfInputs(9)        'P8
    parBundleFlangeDia = arrayOfInputs(10)          'P7
    parBundleFlangeTk = arrayOfInputs(11)           'P6
    parBundlePullingLength = arrayOfInputs(12)      'P9
    parSupportLength = arrayOfInputs(13)            'P15
    parBotSupportCenFromPP = arrayOfInputs(14)      'P10
    parBottomSupportCentoCen = arrayOfInputs(15)    'P11
    parBottomSupportHeight = arrayOfInputs(16)      'P14
    parSupport1Thickness = arrayOfInputs(17)        'P12
    parSupport2Thickness = arrayOfInputs(18)        'P13
    parFrontEndFlangeDia = arrayOfInputs(19)        'P30
    parFrontEndFlangeTk1 = arrayOfInputs(20)        'P31
    parFrontEndFlangeTk2 = arrayOfInputs(21)        'P34
    parFrontEndFlangeTk3 = arrayOfInputs(22)        'P35
    parFrontEndLength1 = arrayOfInputs(23)          'P32
    parFrontEndLength2 = arrayOfInputs(24)          'P33
    parFrontEndDiameter = arrayOfInputs(25)         'P33 of type Q End
    parEndHeadConeTopDiameter = arrayOfInputs(26)   'P51
    parEndHeadConeHeight = arrayOfInputs(27)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(28)     'P51 for TORC & TORS
    parEndHeadDomeradius = arrayOfInputs(29)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(30)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(31)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(32)   'P50 for TORS
    parInsulationThickness = arrayOfInputs(33)
    parFP1XBoltHole = arrayOfInputs(34)
    parFP1YBoltHole = arrayOfInputs(35)
    parFP2XBoltHole = arrayOfInputs(36)
    parFP2YBoltHole = arrayOfInputs(37)

    iOutput = 0
    
    'Insert your code for output 12(Bundle Pulling Cylinder)
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    
    'P32
    startPoint.Set parFrontEndLength1, 0, 0
    'P9-P32
    endPoint.Set (parFrontEndLength1 - parBundlePullingLength), 0, 0
                           
    Set ObjBundlePullCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                parExchangerNeckDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundlePullCylin
    Set ObjBundlePullCylin = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


