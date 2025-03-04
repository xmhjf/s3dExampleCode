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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Friday, Feb 27 2004
'   Description:
'' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'   E310 Symbol in Page no 305 of pdf file.  This symbol has variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'   25.Aug.2008     VKC       TR-110721 Insulation around the flanges is modified.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
PI = 4 * Atn(1)
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjExchanBodyIns As Object
    Dim ObjFrontEndBodyIns As Object
    Dim ObjRearEndBodyIns As Object
    Dim ObjExpanJointIns As Object
    Dim ObjNozzle1Ins As Object
    Dim ObjNozzle2Ins As Object
    Dim ObjNozzle3Ins As Object
    Dim ObjNozzle4Ins As Object
    
    
    Dim parExchangerLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parExchangerFlangeTk2 As Double
    Dim parExtensionJointPosition As Double
    Dim parExtensionJointThickness As Double
    Dim parExtensionJointDiameter As Double
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
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parRearEndFlangeTk2 As Double
    Dim parRearEndFlangeTk3 As Double
    Dim parInsulationThickness As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)           'P1
    parExchangerDiameter = arrayOfInputs(3)         'P2
    parBundleFlangeTk = arrayOfInputs(4)                'P3
    parBundleFlangeDia = arrayOfInputs(5)               'P4
    parExchangerFlangeTk1 = arrayOfInputs(6)        'P5
    parExchangerFlangeTk2 = arrayOfInputs(7)        'P6
    parExtensionJointPosition = arrayOfInputs(8)        'P7
    parExtensionJointThickness = arrayOfInputs(9)           'P8
    parExtensionJointDiameter = arrayOfInputs(10)           'P9
    parBundlePullingLength = arrayOfInputs(11)                  'P10
    parSupportAngularLocation = arrayOfInputs(12)           'P11
    parNumberOfSupports = arrayOfInputs(13)                     'P12
    parSupportDepth = arrayOfInputs(14)                                 'P13
    parSupportFromOrigin = arrayOfInputs(15)                    'P14
    parSupportTopWidth = arrayOfInputs(16)                  'P15
    parSupportBottomWidth = arrayOfInputs(17)               'P16
    parSupportThickness = arrayOfInputs(18)                 'P17
    parFrontEndFlangeDia = arrayOfInputs(19)            'P30
    parFrontEndFlangeTk1 = arrayOfInputs(20)        'P31
    parFrontEndLength1 = arrayOfInputs(21)          'P32
    parFrontEndLength2 = arrayOfInputs(22)          'P33
    parFrontEndFlangeTk2 = arrayOfInputs(23)            'P34
    parFrontEndFlangeTk3 = arrayOfInputs(24)            'P35
    parRearEndFlangeDia = arrayOfInputs(25)             'P40
    parRearEndFlangeTk1 = arrayOfInputs(26)             'P41
    parRearEndLength = arrayOfInputs(27)                    'P42
    parRearEndFlangeTk2 = arrayOfInputs(28)                 'P43
    parRearEndFlangeTk3 = arrayOfInputs(29)                 'P44
    parInsulationThickness = arrayOfInputs(30)
    
    iOutput = 0

    ''Origin is at DP/PP1

   Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
' Insert your code for output 5 (ExchangerBody Insulation)
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
    dblStrPts(5) = -(parFrontEndLength1 + parExchangerLength)
    'Point3
    dblStrPts(6) = (parRearEndFlangeDia / 2) + parInsulationThickness
    dblStrPts(7) = 0
    dblStrPts(8) = -(parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2 _
                                                            - parInsulationThickness)
    'Point4
    dblStrPts(9) = (parExchangerDiameter / 2) + parInsulationThickness
    dblStrPts(10) = 0
    dblStrPts(11) = dblStrPts(8)
    'Point5
    dblStrPts(12) = dblStrPts(9)
    dblStrPts(13) = 0
    dblStrPts(14) = -(parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1 _
                                                            + parInsulationThickness)
    If CmpDblGreaterthan(parBundleFlangeDia, parFrontEndFlangeDia) Then
    'Point6
    dblStrPts(15) = (parBundleFlangeDia / 2) + parInsulationThickness
    dblStrPts(16) = 0
    dblStrPts(17) = dblStrPts(14)
    'Point7
    dblStrPts(18) = (parBundleFlangeDia / 2) + parInsulationThickness
    dblStrPts(19) = 0
    dblStrPts(20) = -(parFrontEndLength1 - parInsulationThickness)
    End If
    If CmpDblGreaterthan(parFrontEndFlangeDia, parBundleFlangeDia) Then
    'Point6
    dblStrPts(15) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblStrPts(16) = 0
    dblStrPts(17) = dblStrPts(14)
    'Point7
    dblStrPts(18) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblStrPts(19) = 0
    dblStrPts(20) = -(parFrontEndLength1 - parInsulationThickness)
    End If
    'Point8
    dblStrPts(21) = 0
    dblStrPts(22) = 0
    dblStrPts(23) = -(parFrontEndLength1 - parInsulationThickness)
    
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
    
 ' Insert your code for output 6(Front End Body Ins)
    Dim oEndLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 26) As Double
    'Point 1
    dblEndPts(0) = 0
    dblEndPts(1) = 0
    dblEndPts(2) = -(parFrontEndLength1 + parInsulationThickness)
    'Point 2
    dblEndPts(3) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(4) = 0
    dblEndPts(5) = -(parFrontEndLength1 + parInsulationThickness)
    'Point 3
    dblEndPts(6) = dblEndPts(3)
    dblEndPts(7) = 0
    dblEndPts(8) = -(parFrontEndLength1 - parFrontEndFlangeTk1 _
                                                            - parInsulationThickness)
    'Point 4
    dblEndPts(9) = (parExchangerDiameter / 2) + parInsulationThickness
    dblEndPts(10) = 0
    dblEndPts(11) = dblEndPts(8)
    'Point 5
    dblEndPts(12) = dblEndPts(9)
    dblEndPts(13) = 0
    dblEndPts(14) = (parFrontEndLength2 - parFrontEndFlangeTk3 - parInsulationThickness)
    'Point 6
    dblEndPts(15) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(16) = 0
    dblEndPts(17) = dblEndPts(14)
'    'Point 7
'    dblEndPts(18) = dblEndPts(15)
'    dblEndPts(19) = 0
'    dblEndPts(20) = parFrontEndLength2
    'Point 8
    dblEndPts(18) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(19) = 0
    dblEndPts(20) = (parFrontEndLength2 + parFrontEndFlangeTk2 + parInsulationThickness)
    'Point 9
    dblEndPts(21) = 0
    dblEndPts(22) = 0
    dblEndPts(23) = dblEndPts(20)
    
    Set oEndLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, (dblEndPts(2) + dblEndPts(26)) / 2

    Set ObjFrontEndBodyIns = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontEndBodyIns
    Set ObjFrontEndBodyIns = Nothing
    
 ' Insert your code for output 7(Rear End Body Ins)
    'Point 1
    dblEndPts(0) = 0
    dblEndPts(1) = 0
    dblEndPts(2) = -(parFrontEndLength1 + parExchangerLength)
    'Point 2
    dblEndPts(3) = (parRearEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(4) = 0
    dblEndPts(5) = dblEndPts(2)
    'Point 3
    dblEndPts(6) = dblEndPts(3)
    dblEndPts(7) = 0
    dblEndPts(8) = -(parFrontEndLength1 + parExchangerLength + parRearEndFlangeTk1 _
                                                       + parInsulationThickness)
    'Point 4
    dblEndPts(9) = (parExchangerDiameter / 2) + parInsulationThickness
    dblEndPts(10) = 0
    dblEndPts(11) = dblEndPts(8)
    'Point 5
    dblEndPts(12) = dblEndPts(9)
    dblEndPts(13) = 0
    dblEndPts(14) = -(parFrontEndLength1 + parExchangerLength + parRearEndLength _
                        - parRearEndFlangeTk3 - parInsulationThickness)
    'Point 6
    dblEndPts(15) = (parRearEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(16) = 0
    dblEndPts(17) = dblEndPts(14)
    'Point 7
    dblEndPts(18) = dblEndPts(15)
    dblEndPts(19) = 0
    dblEndPts(20) = -(parFrontEndLength1 + parExchangerLength + parRearEndLength _
                        + parRearEndFlangeTk2 + parInsulationThickness)
    'Point 8
    dblEndPts(21) = 0
    dblEndPts(22) = 0
    dblEndPts(23) = dblEndPts(20)
    
    Set oEndLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, (dblEndPts(2) + dblEndPts(26)) / 2

    Set ObjRearEndBodyIns = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearEndBodyIns
    Set ObjRearEndBodyIns = Nothing
    
 ' Insert your code for output 8(Extension Joint Ins)
    'P7 and P8 are defined provide Extension Joint
    If (CmpDblGreaterthan(parExtensionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExtensionJointThickness, LINEAR_TOLERANCE)) Then
        Dim ObjExpenJtIns As Object
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set 0, 0, -(parExtensionJointPosition - parInsulationThickness)
        endPoint.Set 0, 0, -(parExtensionJointPosition + parExtensionJointThickness + parInsulationThickness)

        Set ObjExpenJtIns = PlaceCylinder(m_OutputColl, startPoint, _
                                            endPoint, (parExtensionJointDiameter + parInsulationThickness * 2), True)
    
    ' Set the output
        m_OutputColl.AddOutput "ExtensionJt_", ObjExpenJtIns
        Set ObjExpenJtIns = Nothing
        Set startPoint = Nothing
        Set endPoint = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
       Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
