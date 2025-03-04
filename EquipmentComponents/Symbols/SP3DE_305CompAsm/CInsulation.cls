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
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Wednesday, Feb 25 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'    E305 Symbol in Page no 299.  Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   29.Nov.2004     V6UpgradeSO     Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc             DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl             TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'   17.Sep.2007     PK              TR-117405 Modified code to ensure that insulation does not extend to
'                                   flanges when P34,P35,P43 and P44 (flange thicknesses) are made zero.
'   25.Aug.2008     VKC             TR-110721 Insulation around the flanges is  modified.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
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
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)                       'P1
    parExchangerDiameter = arrayOfInputs(3)                     'P2
    parBundleFlangeTk = arrayOfInputs(4)                            'P3
    parBundleFlangeDia = arrayOfInputs(5)                           'P4
    parExchangerFlangeTk1 = arrayOfInputs(6)                    'P5
    parExchangerFlangeTk2 = arrayOfInputs(7)                    'P6
    parExpansionJointPosition = arrayOfInputs(8)                'P7
    parExpansionJointThickness = arrayOfInputs(9)               'P8
    parExpansionJointDiameter = arrayOfInputs(10)               'P9
    parBundlePullingLength = arrayOfInputs(11)                  'P10
    parBotSupportCenFromPP = arrayOfInputs(12)              'P11
    parBottomSupportCentoCen = arrayOfInputs(13)            'P12
    parSupport1Thickness = arrayOfInputs(14)                        'P13
    parSupport2Thickness = arrayOfInputs(15)                        'P14
    parBottomSupportHeight = arrayOfInputs(16)                  'P15
    parSupportLength = arrayOfInputs(17)                                'P16
    parTopSupportCenFromPP = arrayOfInputs(18)                  'P17
    parTopSupportCentoCen = arrayOfInputs(19)                   'P18
    parTopSupportHeight = arrayOfInputs(20)                         'P19
    parFrontEndFlangeDia = arrayOfInputs(21)                    'P30
    parFrontEndFlangeTk1 = arrayOfInputs(22)                    'P31
    parFrontEndLength1 = arrayOfInputs(23)          'P32
    parFrontEndLength2 = arrayOfInputs(24)          'P33
    parFrontEndFlangeTk2 = arrayOfInputs(25)      'P34
    parFrontEndFlangeTk3 = arrayOfInputs(26)      'P35
    parRearEndFlangeDia = arrayOfInputs(27)        'P40
    parRearEndFlangeTk1 = arrayOfInputs(28)         'P41
    parRearEndLength = arrayOfInputs(29)                'P42
    parRearEndFlangeTk2 = arrayOfInputs(30)             'P43
    parRearEndFlangeTk3 = arrayOfInputs(31)            'P44
    parInsulationThickness = arrayOfInputs(32)
    
    iOutput = 0
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    If CmpDblEqual(parBundlePullingLength, LINEAR_TOLERANCE) Then parBundlePullingLength = parExchangerLength
 ' Insert your code for output 9(Exchanger Insullation)
    'Form a line string and Revolve it about X-Axis
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 29) As Double
    'Point 1 Exchanger Flange at Rear End
    dblStrPts(0) = (parFrontEndLength1 + parExchangerLength + parRearEndFlangeTk1 _
                    + parInsulationThickness)
    dblStrPts(1) = CenterPos.y
    dblStrPts(2) = CenterPos.z
    'Point2
    dblStrPts(3) = dblStrPts(0)
    dblStrPts(4) = CenterPos.y
    dblStrPts(5) = parRearEndFlangeDia / 2 + parInsulationThickness
    'Point3
    dblStrPts(6) = (dblStrPts(3)) - (parExchangerFlangeTk2) - parRearEndFlangeTk1 _
    - 2 * parInsulationThickness
    dblStrPts(7) = CenterPos.y
    dblStrPts(8) = parRearEndFlangeDia / 2 + parInsulationThickness
    'Point4
    dblStrPts(9) = dblStrPts(6)
    dblStrPts(10) = CenterPos.y
    dblStrPts(11) = parExchangerDiameter / 2 + parInsulationThickness
    'Point5
    dblStrPts(12) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1 + parInsulationThickness)
    dblStrPts(13) = CenterPos.y
    dblStrPts(14) = parExchangerDiameter / 2 + parInsulationThickness
    If CmpDblGreaterthan(parRearEndFlangeDia, parBundleFlangeDia) And CmpDblGreaterthan(parRearEndFlangeDia, parFrontEndFlangeDia) Then
    'Point6
    dblStrPts(15) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1 + parInsulationThickness)
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = parRearEndFlangeDia / 2 + parInsulationThickness
    'Point7
    dblStrPts(18) = parFrontEndLength1 - parFrontEndFlangeTk1 _
    - parInsulationThickness
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = parRearEndFlangeDia / 2 + parInsulationThickness
    End If
    If CmpDblGreaterthan(parBundleFlangeDia, parRearEndFlangeDia) And CmpDblGreaterthan(parBundleFlangeDia, parFrontEndFlangeDia) Then
    'Point6
    dblStrPts(15) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1 + parInsulationThickness)
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = parBundleFlangeDia / 2 + parInsulationThickness
    'Point7
    dblStrPts(18) = parFrontEndLength1 - parFrontEndFlangeTk1 - parInsulationThickness
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = parBundleFlangeDia / 2 + parInsulationThickness
    End If
    If CmpDblGreaterthan(parFrontEndFlangeDia, parBundleFlangeDia) And CmpDblGreaterthan(parFrontEndFlangeDia, parRearEndFlangeDia) Then
    'Point6
    dblStrPts(15) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1 + parInsulationThickness)
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = parFrontEndFlangeDia / 2 + parInsulationThickness
    'Point7
    dblStrPts(18) = parFrontEndLength1 - parFrontEndFlangeTk1 - parInsulationThickness
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = parFrontEndFlangeDia / 2 + parInsulationThickness
    End If
    'Point8
    dblStrPts(21) = parFrontEndLength1 - parFrontEndFlangeTk1 _
     - parInsulationThickness
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
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblStrPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (parFrontEndLength1 + parExchangerLength / 2), _
                                    CenterPos.y, _
                                    CenterPos.z                     'P32+P1 /2
    
    Dim ObjExchangerInsul As Object
    
    Set ObjExchangerInsul = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerInsul
    Set ObjExchangerInsul = Nothing
    Set oLineStr = Nothing
    
'Expansion Joint Insulation
    'P7 and P8 are defined provide expansion Joint
    If CmpDblGreaterthan(parExpansionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExpansionJointThickness, LINEAR_TOLERANCE) Then
        Dim ObjExpanJt  As Object
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set parExpansionJointPosition - parInsulationThickness, CenterPos.y, CenterPos.z
        endPoint.Set parExpansionJointPosition + parExpansionJointThickness + parInsulationThickness, CenterPos.y, CenterPos.z
               
        Set ObjExpanJt = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                        parExpansionJointDiameter + parInsulationThickness * 2, True)
        
 ' Set the output
        m_OutputColl.AddOutput "ExpansionJtIns_", ObjExpanJt
        Set ObjExpanJt = Nothing
    End If
    
  ' Insert your code for output 6(FrontEndBody)
    'Create a Point String and Revolve it
    Dim oEndLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 26) As Double
    
    'Point 1
    dblEndPts(0) = parFrontEndLength1        '+P32
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
    dblEndPts(12) = -(parFrontEndLength2 - parFrontEndFlangeTk3 - parInsulationThickness)
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = parExchangerDiameter / 2 + parInsulationThickness
    'Point 6
    dblEndPts(15) = -(parFrontEndLength2 + parFrontEndFlangeTk2 + _
    parInsulationThickness)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parExchangerDiameter / 2 + parInsulationThickness
    
    
    'The following If condition is placed to ensure that insulation does not get extended to flanges
    ' when flange thicknesses P34 and P35 are zero.iPIndex is used to keep track of the coordinates.
    'iPointCount is used to keep track of the number of points used for the linestring
    Dim iPIndex As Integer
    Dim iPointCount As Integer
    If CmpDblLessThanOrEqualTo(parFrontEndFlangeTk2, 0) And CmpDblLessThanOrEqualTo(parFrontEndFlangeTk3, 0) Then
        iPIndex = 12
        iPointCount = 5
    Else
        'Point 6
        dblEndPts(15) = dblEndPts(12)
        dblEndPts(16) = CenterPos.y
        dblEndPts(17) = parFrontEndFlangeDia / 2 + parInsulationThickness
        'Point 7
        dblEndPts(18) = -parFrontEndLength2
        dblEndPts(19) = CenterPos.y
        dblEndPts(20) = parFrontEndFlangeDia / 2 + parInsulationThickness
        'Point 8 changed 21
        dblEndPts(21) = -(parFrontEndLength2 + parFrontEndFlangeTk2 + parInsulationThickness)
        dblEndPts(22) = CenterPos.y
        dblEndPts(23) = parFrontEndFlangeDia / 2 + parInsulationThickness
        iPIndex = 21
        iPointCount = 9
    End If
    
    'Point 9
    dblEndPts(iPIndex + 3) = dblEndPts(iPIndex)
    dblEndPts(iPIndex + 4) = CenterPos.y
    dblEndPts(iPIndex + 5) = CenterPos.z
        
    Set oEndLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, iPointCount, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set 0, 0, 0

    Dim ObjFrontEndInsul As Object
    Set ObjFrontEndInsul = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontEndInsul
    Set ObjFrontEndInsul = Nothing
    
 ' Insert your code for output 11(Rear End Insulation)
    'Point 1 Start Point P32+P1
    Dim stPoint As Double
    stPoint = parExchangerLength + parFrontEndLength1
    dblEndPts(0) = stPoint
    dblEndPts(1) = CenterPos.y
    dblEndPts(2) = CenterPos.z
    'Point 2
    dblEndPts(3) = dblEndPts(0)
    dblEndPts(4) = dblEndPts(1)
    dblEndPts(5) = parRearEndFlangeDia / 2 + parInsulationThickness   'P40/2
    'Point 3
    dblEndPts(6) = (stPoint + parRearEndFlangeTk1 + parInsulationThickness)                  'P41
    dblEndPts(7) = CenterPos.y
    dblEndPts(8) = parRearEndFlangeDia / 2 + parInsulationThickness     'P40/2
    'Point 4
    dblEndPts(9) = dblEndPts(6)
    dblEndPts(10) = CenterPos.y
    dblEndPts(11) = parExchangerDiameter / 2 + parInsulationThickness                 '(P40-P2)/2
    'Point 5
    dblEndPts(12) = (stPoint + parRearEndLength - parRearEndFlangeTk2 _
    - parInsulationThickness)       'P42-P44
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = parExchangerDiameter / 2 + parInsulationThickness
    'Point 6
    dblEndPts(15) = dblEndPts(12)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parRearEndFlangeDia / 2 + parInsulationThickness
    'Point 7
    dblEndPts(18) = stPoint + parRearEndLength + parRearEndFlangeTk3 _
    + parInsulationThickness
    dblEndPts(19) = CenterPos.y
    dblEndPts(20) = parRearEndFlangeDia / 2 + parInsulationThickness
    'Point 8
    dblEndPts(21) = stPoint + parRearEndLength + parRearEndFlangeTk3 _
    + parInsulationThickness
    dblEndPts(22) = CenterPos.y
    dblEndPts(23) = CenterPos.z
    
    'The following If condition is placed to ensure that insulation does not get extended to flanges
    ' when flange thicknesses P43 and P44 are zero.iPIndex is used to keep track of the coordinates.
    'iPointCount is used to keep track of the number of points used for the linestring
    If CmpDblLessThanOrEqualTo(parRearEndFlangeTk2, 0) And CmpDblLessThanOrEqualTo(parRearEndFlangeTk3, 0) Then
        iPIndex = 12
        iPointCount = 5
    Else
        'Point 6
        dblEndPts(15) = dblEndPts(12)
        dblEndPts(16) = CenterPos.y
        dblEndPts(17) = parRearEndFlangeDia / 2 + parInsulationThickness
        'Point 7
        dblEndPts(18) = (stPoint + parRearEndLength + parInsulationThickness + parRearEndFlangeTk3)
        dblEndPts(19) = CenterPos.y
        dblEndPts(20) = parRearEndFlangeDia / 2 + parInsulationThickness
        'Point 8 changed
        dblEndPts(21) = (stPoint + parRearEndLength + parInsulationThickness + parRearEndFlangeTk3)          'P43
        dblEndPts(22) = CenterPos.y
        dblEndPts(23) = CenterPos.z
        iPIndex = 21
        iPointCount = 9
    End If
    
    'Point 9
    dblEndPts(iPIndex + 3) = dblEndPts(iPIndex)
    dblEndPts(iPIndex + 4) = dblEndPts(iPIndex + 1)
    dblEndPts(iPIndex + 5) = CenterPos.z

    Set oEndLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, iPointCount, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (stPoint + parRearEndLength / 2), CenterPos.y, CenterPos.z

    Dim ObjRearEndInsul As Object
    Set ObjRearEndInsul = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearEndInsul
    Set ObjRearEndInsul = Nothing
    Set oGeomFactory = Nothing
    
    Set revVector = Nothing
    Set revCenterPt = Nothing
    Set oEndLineStr = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set CenterPos = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
