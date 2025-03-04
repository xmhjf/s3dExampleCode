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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:38 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author: VRK
'   Creation Date:  Monday, Feb 26 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Four Outputs:
'         In which Insulated ExchangerCylinder, FrontEnd, RearEnd, and Expansionjoints are created.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()

    PI = 4 * Atn(1)

End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim CenterPos As New AutoMath.DPosition
    Dim iOutput As Double
    Dim ObjExchangerInsul As Object
    Dim ObjFrontEndInsul As Object
    
    Dim parFrontEndType As Long
    Dim parRearEndType As Long
    Dim parExchangerLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parExchangerFlangeTk2 As Double
    Dim parExpansionJointPosition As Double
    Dim parExpansionJointThickness As Double
    Dim parExpansionJointDiameter As Double
    Dim parBundlePullingLength As Double
    Dim parBotSupportCenFromPP As Double
    Dim parBottomSupportCentoCen As Double
    Dim parSupport1Thickness As Double
    Dim parSupport2Thickness As Double
    Dim parBottomSupportHeight As Double
    Dim parSupportLength As Double
    Dim parTopSupportCenFromPP As Double
    Dim parTopSupportCentoCen As Double
    Dim parTopSupportHeight As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parRearEndFlangeTk2 As Double
    Dim parRearEndFlangeTk3 As Double
    Dim parChannelDiameter As Double
    Dim parFrontEndDiameter As Double
    Dim parRearEndDiameter As Double
    Dim parInsulationThickness As Double
    Dim parFP1XBoltHole As Double
    Dim parFP1YBoltHole As Double
    Dim parFP2XBoltHole As Double
    Dim parFP2YBoltHole As Double
    
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndType = arrayOfInputs(2)
    parRearEndType = arrayOfInputs(3)
    parExchangerDiameter = arrayOfInputs(4)                     'P2
    parExchangerLength = arrayOfInputs(5)                       'P1
    parExchangerFlangeTk1 = arrayOfInputs(6)                    'P5
    parExchangerFlangeTk2 = arrayOfInputs(7)                    'P6
    parExpansionJointDiameter = arrayOfInputs(8)                'P9
    parExpansionJointPosition = arrayOfInputs(9)                'P7
    parExpansionJointThickness = arrayOfInputs(10)              'P8
    parBundleFlangeDia = arrayOfInputs(11)                      'P4
    parBundleFlangeTk = arrayOfInputs(12)                       'P3
    parBundlePullingLength = arrayOfInputs(13)                  'P10
    parSupportLength = arrayOfInputs(14)                        'P16
    parBotSupportCenFromPP = arrayOfInputs(15)                  'P11
    parBottomSupportCentoCen = arrayOfInputs(16)                'P12
    parBottomSupportHeight = arrayOfInputs(17)                  'P15
    parSupport1Thickness = arrayOfInputs(18)                    'P13
    parSupport2Thickness = arrayOfInputs(19)                    'P14
    parTopSupportHeight = arrayOfInputs(20)                     'P19
    parTopSupportCenFromPP = arrayOfInputs(21)                  'P17
    parTopSupportCentoCen = arrayOfInputs(22)                   'P18
    parFrontEndFlangeDia = arrayOfInputs(23)                    'P30
    parFrontEndFlangeTk1 = arrayOfInputs(24)                    'P31
    parFrontEndFlangeTk2 = arrayOfInputs(25)                    'P34
    parFrontEndFlangeTk3 = arrayOfInputs(26)                    'P35
    parFrontEndLength1 = arrayOfInputs(27)                      'P32
    parFrontEndLength2 = arrayOfInputs(28)                      'P33
    parRearEndFlangeDia = arrayOfInputs(29)                     'P40
    parRearEndFlangeTk1 = arrayOfInputs(30)                     'P41
    parRearEndFlangeTk2 = arrayOfInputs(31)                     'P43
    parRearEndFlangeTk3 = arrayOfInputs(32)                     'P44
    parRearEndLength = arrayOfInputs(33)                        'P42
    parChannelDiameter = arrayOfInputs(34)                      'P43 of type M/S/T/U/W2 End
    parFrontEndDiameter = arrayOfInputs(35)                     'P33 of type Q End
    parRearEndDiameter = arrayOfInputs(36)                      'P43 of type Q End
    parInsulationThickness = arrayOfInputs(37)
    parFP1XBoltHole = arrayOfInputs(38)
    parFP1YBoltHole = arrayOfInputs(39)
    parFP2XBoltHole = arrayOfInputs(40)
    parFP2YBoltHole = arrayOfInputs(41)
    
    iOutput = 0
    CenterPos.Set 0, 0, 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    If CmpDblEqual(parBundlePullingLength, LINEAR_TOLERANCE) Then parBundlePullingLength = parExchangerLength
    'Insert your code for Exchanger Insulation
    'Form a line string and Revolve it about X-Axis
    Dim oLinestr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 29) As Double
    
    'Point 1 Exchanger Flange at Rear End
    dblStrPts(0) = (parFrontEndLength1 + parExchangerLength)
    dblStrPts(1) = CenterPos.y
    dblStrPts(2) = CenterPos.z
    
    'Point2
    dblStrPts(3) = dblStrPts(0)
    dblStrPts(4) = CenterPos.y
    dblStrPts(5) = parRearEndFlangeDia / 2 + parInsulationThickness
    
    'Point3
    dblStrPts(6) = (dblStrPts(3)) - (parExchangerFlangeTk2)
    dblStrPts(7) = CenterPos.y
    dblStrPts(8) = parRearEndFlangeDia / 2 + parInsulationThickness
    
    'Point4
    dblStrPts(9) = dblStrPts(6)
    dblStrPts(10) = CenterPos.y
    dblStrPts(11) = parExchangerDiameter / 2 + parInsulationThickness
    
    'Point5
    dblStrPts(12) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    dblStrPts(13) = CenterPos.y
    dblStrPts(14) = parExchangerDiameter / 2 + parInsulationThickness
    
    'Point6
    dblStrPts(15) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = parFrontEndFlangeDia / 2 + parInsulationThickness
    
    'Point7
    dblStrPts(18) = (parFrontEndLength1 + parBundleFlangeTk)
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = parFrontEndFlangeDia / 2 + parInsulationThickness
    
    'Point8
    dblStrPts(21) = (parFrontEndLength1 + parBundleFlangeTk)
    dblStrPts(22) = CenterPos.y
    dblStrPts(23) = parBundleFlangeDia / 2 + parInsulationThickness
    
    'Point9
    dblStrPts(24) = (parFrontEndLength1)
    dblStrPts(25) = CenterPos.y
    dblStrPts(26) = parBundleFlangeDia / 2 + parInsulationThickness
    
    'Point10
    dblStrPts(27) = (parFrontEndLength1)
    dblStrPts(28) = CenterPos.y
    dblStrPts(29) = CenterPos.z
    
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition
    Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblStrPts)
    
    revVector.Set 1, 0, 0
    revCenterPt.Set (parFrontEndLength1 + parExchangerLength / 2), _
                                    CenterPos.y, _
                                    CenterPos.z                     'P32+P1 /2
    
    Set ObjExchangerInsul = PlaceRevolution(m_outputColl, oLinestr, revVector, revCenterPt, PI * 2, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerInsul
    Set ObjExchangerInsul = Nothing
    Set oLinestr = Nothing
   
    'Insert your code for FrontEndBody
    'Create a Point String and Revolve it
    Dim oEndLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 26) As Double
    
    'Point 1
    dblEndPts(0) = parFrontEndLength1        'P32
    dblEndPts(1) = CenterPos.y
    dblEndPts(2) = CenterPos.z
    
    'Point 2
    dblEndPts(3) = parFrontEndLength1
    dblEndPts(4) = CenterPos.y
    dblEndPts(5) = parFrontEndFlangeDia / 2 + parInsulationThickness   'P30/2
    
    'Point 3
    dblEndPts(6) = (parFrontEndLength1 - parFrontEndFlangeTk1)
    dblEndPts(7) = CenterPos.y
    dblEndPts(8) = parFrontEndFlangeDia / 2 + parInsulationThickness  'P30/2
    
    'Point 4
    dblEndPts(9) = dblEndPts(6)
    dblEndPts(10) = dblEndPts(7)
    dblEndPts(11) = parExchangerDiameter / 2 + parInsulationThickness
    
    'Point 5
    dblEndPts(12) = -(parFrontEndLength2 - parFrontEndFlangeTk3)
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = parExchangerDiameter / 2 + parInsulationThickness
    
    'Point 6
    dblEndPts(15) = dblEndPts(12)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parFrontEndFlangeDia / 2 + parInsulationThickness
    
    'Point 7
    dblEndPts(18) = -parFrontEndLength2
    dblEndPts(19) = CenterPos.y
    dblEndPts(20) = parFrontEndFlangeDia / 2 + parInsulationThickness
    
    'Point 8
    dblEndPts(21) = -(parFrontEndLength2 + parFrontEndFlangeTk2)
    dblEndPts(22) = CenterPos.y
    dblEndPts(23) = parFrontEndFlangeDia / 2 + parInsulationThickness
    
    'Point 9
    dblEndPts(24) = dblEndPts(21)
    dblEndPts(25) = CenterPos.y
    dblEndPts(26) = CenterPos.z
    
    Set oEndLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set 0, 0, 0
    Set ObjFrontEndInsul = PlaceRevolution(m_outputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontEndInsul
    Set ObjFrontEndInsul = Nothing
    
    'Insert your code for Rear End Insulation
    Dim stPoint As Double
    stPoint = parExchangerLength + parFrontEndLength1
    'Point 1 Start Point P32+P1
    dblEndPts(0) = stPoint
    dblEndPts(1) = CenterPos.y
    dblEndPts(2) = CenterPos.z
    
    'Point 2
    dblEndPts(3) = dblEndPts(0)
    dblEndPts(4) = dblEndPts(1)
    dblEndPts(5) = parRearEndFlangeDia / 2 + parInsulationThickness   'P40/2
    
    'Point 3
    dblEndPts(6) = (stPoint + parRearEndFlangeTk1)                  'P41
    dblEndPts(7) = CenterPos.y
    dblEndPts(8) = parRearEndFlangeDia / 2 + parInsulationThickness     'P40/2
    
    'Point 4
    dblEndPts(9) = dblEndPts(6)
    dblEndPts(10) = CenterPos.y
    dblEndPts(11) = parExchangerDiameter / 2 + parInsulationThickness                 '(P40-P2)/2
    
    'Point 5
    dblEndPts(12) = (stPoint + parRearEndLength - parRearEndFlangeTk3)       'P42-P44
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = parExchangerDiameter / 2 + parInsulationThickness
    
    'Point 6
    dblEndPts(15) = dblEndPts(12)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parRearEndFlangeDia / 2 + parInsulationThickness
    
    'Point 7
    dblEndPts(18) = (stPoint + parRearEndLength)
    dblEndPts(19) = CenterPos.y
    dblEndPts(20) = parRearEndFlangeDia / 2 + parInsulationThickness
    
    'Point 8
    dblEndPts(21) = (stPoint + parRearEndLength + parRearEndFlangeTk2)              'P43
    dblEndPts(22) = CenterPos.y
    dblEndPts(23) = parRearEndFlangeDia / 2 + parInsulationThickness
    
    'Point 9
    dblEndPts(24) = dblEndPts(21)
    dblEndPts(25) = dblEndPts(22)
    dblEndPts(26) = CenterPos.z

    Set oEndLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (stPoint + parRearEndLength / 2), CenterPos.y, CenterPos.z

    Dim ObjRearEndInsul As Object
    Set ObjRearEndInsul = PlaceRevolution(m_outputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjRearEndInsul
    Set ObjRearEndInsul = Nothing
    Set oGeomFactory = Nothing
    Set revVector = Nothing
    Set revCenterPt = Nothing
    Set oEndLineStr = Nothing
     
    'Expansion Joint Insulation
    'P7 and P8 are defined provide expansion Joint
    If CmpDblGreaterthan(parExpansionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExpansionJointThickness, LINEAR_TOLERANCE) Then
        Dim ObjExpanJt  As Object
        startPoint.Set parExpansionJointPosition, CenterPos.y, CenterPos.z
        endPoint.Set parExpansionJointPosition + parExpansionJointThickness, CenterPos.y, CenterPos.z

        Set ObjExpanJt = PlaceCylinder(m_outputColl, startPoint, endPoint, _
                                                        parExpansionJointDiameter + parInsulationThickness * 2, True)

        'Set the output
        m_outputColl.AddOutput "ExpansionJtIns_", ObjExpanJt
        Set ObjExpanJt = Nothing

    End If
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set CenterPos = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
