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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:45 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MS
'   Creation Date:  Thursday, May 13 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol detail is based on the PDS Equipment Modeling User's Guide,
'    E305 Symbol in Page no 299.  Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'   22.Aug.2008     GL              TR-110721 Modified the code to ensure that insulation is proper
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
    Dim parChannelDiameter As Double
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
    Dim parFrontEndLength1 As Double
    Dim parRearEndFlangeDia As Double
        
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)                       'P1
    parChannelDiameter = arrayOfInputs(3)                     'P2
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
    parFrontEndLength1 = arrayOfInputs(22)          'P32
    parRearEndFlangeDia = arrayOfInputs(23)        'P40
    parInsulationThickness = arrayOfInputs(24)
    
    iOutput = 0
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    If CmpDblEqual(parBundlePullingLength, LINEAR_TOLERANCE) Then parBundlePullingLength = parExchangerLength
 ' Insert your code for output 5 (Exchanger Insulation)
    'Form a line string and Revolve it about X-Axis
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 29) As Double
    'Point 1 Exchanger Flange at Rear End
    dblStrPts(0) = (parFrontEndLength1 + parExchangerLength + parInsulationThickness)
    dblStrPts(1) = CenterPos.y
    dblStrPts(2) = CenterPos.z
    'Point2
    dblStrPts(3) = dblStrPts(0)
    dblStrPts(4) = CenterPos.y
    dblStrPts(5) = parRearEndFlangeDia / 2 + parInsulationThickness
      'Point3
    dblStrPts(6) = parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2 - parInsulationThickness
    dblStrPts(7) = CenterPos.y
    dblStrPts(8) = parRearEndFlangeDia / 2 + parInsulationThickness
    'Point4
    dblStrPts(9) = dblStrPts(6)
    dblStrPts(10) = CenterPos.y
    dblStrPts(11) = parChannelDiameter / 2 + parInsulationThickness
    'Point5''
    dblStrPts(12) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1) + parInsulationThickness
    dblStrPts(13) = CenterPos.y
    dblStrPts(14) = parChannelDiameter / 2 + parInsulationThickness
    'Point6
    dblStrPts(15) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1) + parInsulationThickness
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = parFrontEndFlangeDia / 2 + parInsulationThickness
    'Point7
    dblStrPts(18) = (parFrontEndLength1 + parBundleFlangeTk) + parInsulationThickness
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = parFrontEndFlangeDia / 2 + parInsulationThickness
    'Point8
    dblStrPts(21) = (parFrontEndLength1 + parBundleFlangeTk) + parInsulationThickness
    dblStrPts(22) = CenterPos.y
    dblStrPts(23) = parBundleFlangeDia / 2 + parInsulationThickness
    If CmpDblGreaterthan(parFrontEndFlangeDia, parBundleFlangeDia) Then
        'Point7
        dblStrPts(18) = (parFrontEndLength1 + parBundleFlangeTk) - parInsulationThickness
        dblStrPts(19) = CenterPos.y
        dblStrPts(20) = parFrontEndFlangeDia / 2 + parInsulationThickness
        'Point8
        dblStrPts(21) = (parFrontEndLength1 + parBundleFlangeTk) - parInsulationThickness
        dblStrPts(22) = CenterPos.y
        dblStrPts(23) = parBundleFlangeDia / 2 + parInsulationThickness
    End If
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
    
    Set oGeomFactory = Nothing
    
    Set revVector = Nothing
    Set revCenterPt = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set CenterPos = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
