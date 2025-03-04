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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Monday, Mar 5 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Two Outputs:
'    Exhanger Body and Expansion Joint
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim ObjExchanBodyIns    As Object
    Dim ObjExpanJtIns    As Object
    Dim iOutput             As Double
    
    Dim parFrontEndType As Double
    Dim parRearEndType As Double
    Dim parExchangerDiameter As Double
    Dim parExchangerLength As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parExchangerFlangeTk2 As Double
    Dim parExpansionJointDiameter As Double
    Dim parExpansionJointPosition As Double
    Dim parExpansionJointThickness As Double
    Dim parBundleFlangeDia As Double
    Dim parBundleFlangeTk As Double
    Dim parBundlePullingLength As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports As Double
    Dim parSupportDepth As Double
    Dim parSupportFromOrigin As Double
    Dim parSupportTopWidth As Double
    Dim parSupportBottomWidth As Double
    Dim parSupportThickness As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndFlangeTk2 As Double
    Dim parRearEndFlangeTk3 As Double
    Dim parRearEndLength As Double
    Dim parChannelDiameter As Double
    Dim parFrontEndDiameter As Double
    Dim parRearEndDiameter As Double
    Dim parInsulationThickness As Double
    
    parFrontEndType = arrayOfInputs(2)
    parRearEndType = arrayOfInputs(3)
    parExchangerDiameter = arrayOfInputs(4)         'P2
    parExchangerLength = arrayOfInputs(5)           'P1
    parExchangerFlangeTk1 = arrayOfInputs(6)        'P5
    parExchangerFlangeTk2 = arrayOfInputs(7)        'P6
    parExpansionJointDiameter = arrayOfInputs(8)    'P9
    parExpansionJointPosition = arrayOfInputs(9)    'P7
    parExpansionJointThickness = arrayOfInputs(10)  'P8
    parBundleFlangeDia = arrayOfInputs(11)          'P4
    parBundleFlangeTk = arrayOfInputs(12)           'P3
    parBundlePullingLength = arrayOfInputs(13)      'P10
    parSupportAngularLocation = arrayOfInputs(14)   'P11
    parNumberOfSupports = arrayOfInputs(15)         'P12
    parSupportDepth = arrayOfInputs(16)             'P13
    parSupportFromOrigin = arrayOfInputs(17)        'P14
    parSupportTopWidth = arrayOfInputs(18)          'P15
    parSupportBottomWidth = arrayOfInputs(19)       'P16
    parSupportThickness = arrayOfInputs(20)         'P17
    parFrontEndFlangeDia = arrayOfInputs(21)        'P30
    parFrontEndFlangeTk1 = arrayOfInputs(22)        'P31
    parFrontEndFlangeTk2 = arrayOfInputs(23)        'P34
    parFrontEndFlangeTk3 = arrayOfInputs(24)        'P35
    parFrontEndLength1 = arrayOfInputs(25)          'P32
    parFrontEndLength2 = arrayOfInputs(26)          'P33
    parRearEndFlangeDia = arrayOfInputs(27)         'P40
    parRearEndFlangeTk1 = arrayOfInputs(28)         'P41
    parRearEndFlangeTk2 = arrayOfInputs(29)         'P43
    parRearEndFlangeTk3 = arrayOfInputs(30)         'P44
    parRearEndLength = arrayOfInputs(31)            'P42
    parChannelDiameter = arrayOfInputs(32)          'P43 of type M/S/T/U/W2 End
    parFrontEndDiameter = arrayOfInputs(33)         'P33 of type Q End
    parRearEndDiameter = arrayOfInputs(34)          'P43 of type Q End
    parInsulationThickness = arrayOfInputs(35)
    
    iOutput = 0

    'Origin is at DP/PP1

    Dim geomFactory   As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for ExchangerBody Insulation
    'Form a line string and Revolve it about X-Axis
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 29) As Double
    
    'Point 1
    dblStrPts(0) = 0
    dblStrPts(1) = 0
    dblStrPts(2) = -(parFrontEndLength1 + parExchangerLength)                     '-(P32+P1)
    'Point2
    dblStrPts(3) = (parRearEndFlangeDia / 2) + parInsulationThickness
    dblStrPts(4) = 0
    dblStrPts(5) = dblStrPts(2)
    'Point3
    dblStrPts(6) = dblStrPts(3)
    dblStrPts(7) = 0
    dblStrPts(8) = -(parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2)
    'Point4
    dblStrPts(9) = (parExchangerDiameter / 2) + parInsulationThickness
    dblStrPts(10) = 0
    dblStrPts(11) = dblStrPts(8)
    'Point5
    dblStrPts(12) = dblStrPts(9)
    dblStrPts(13) = 0
    dblStrPts(14) = -(parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    'Point6
    dblStrPts(15) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblStrPts(16) = 0
    dblStrPts(17) = dblStrPts(14)
    'Point7
    dblStrPts(18) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblStrPts(19) = 0
    dblStrPts(20) = -(parFrontEndLength1 + parBundleFlangeTk)
    'Point8
    dblStrPts(21) = (parBundleFlangeDia / 2) + parInsulationThickness
    dblStrPts(22) = 0
    dblStrPts(23) = dblStrPts(20)
    'Point9
    dblStrPts(24) = dblStrPts(21)
    dblStrPts(25) = 0
    dblStrPts(26) = -parFrontEndLength1
    'Point10
    dblStrPts(27) = 0
    dblStrPts(28) = 0
    dblStrPts(29) = -parFrontEndLength1
    
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition

    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblStrPts)
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, (dblStrPts(2) + dblStrPts(29)) / 2
    Set ObjExchanBodyIns = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchanBodyIns
    Set ObjExchanBodyIns = Nothing
    Set oLineStr = Nothing
    
    Set geomFactory = Nothing
    Set revVector = Nothing
    Set revCenterPt = Nothing
    
    'Insert your code for Expansion Joint Ins
    'P7 and P8 are defined provide Expansion Joint
    If (CmpDblGreaterthan(parExpansionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExpansionJointThickness, LINEAR_TOLERANCE)) Then
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set 0, 0, -parExpansionJointPosition - parInsulationThickness
        endPoint.Set 0, 0, -parExpansionJointPosition - parExpansionJointThickness - parInsulationThickness

        Set ObjExpanJtIns = PlaceCylinder(m_OutputColl, startPoint, _
                                            endPoint, (parExpansionJointDiameter + parInsulationThickness * 2), True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExpanJtIns
        Set ObjExpanJtIns = Nothing
        Set startPoint = Nothing
        Set endPoint = Nothing
    End If
    
    Exit Sub
ErrorLabel:
     Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
     Err.HelpFile, Err.HelpContext
End Sub
