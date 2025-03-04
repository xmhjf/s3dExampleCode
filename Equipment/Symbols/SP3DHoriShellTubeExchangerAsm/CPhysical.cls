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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:23 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Monday, Apr 15 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is E305 Horizontal Shell and Tube Exchanger symbol.
'Physical Aspect consist of Exchanger Body, Front End, Rear End and Four Supports

'Insulation Aspect consist of Cylinder for the Exchanger, Cylinder Rear End and Cylinder Front
'End.
'Maintenance Aspect Consist of Cylinder for Bundle Pulling area.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjExchangerBody As Object
    Dim ObjSupport1 As Object
    Dim ObjSupport2 As Object
    Dim ObjSupport3 As Object
    Dim ObjSupport4 As Object
    Dim ObjFrontEndBody As Object
    Dim ObjRearEndBody As Object
    
    
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

    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    If parBundlePullingLength = 0 Then parBundlePullingLength = parExchangerLength

 ' Insert your code for output 1(ExchangerBody)
    'Form a line string and Revolve it about X-Axis
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 29) As Double
    'Point 1 Exchanger Flange at Rear End
    dblStrPts(0) = (parFrontEndLength1 + parExchangerLength)
    dblStrPts(1) = CenterPos.y
    dblStrPts(2) = CenterPos.z
    'Point2
    dblStrPts(3) = dblStrPts(0)
    dblStrPts(4) = CenterPos.y
    dblStrPts(5) = parRearEndFlangeDia / 2
    'Point3
    dblStrPts(6) = (dblStrPts(3)) - (parExchangerFlangeTk2)
    dblStrPts(7) = CenterPos.y
    dblStrPts(8) = parRearEndFlangeDia / 2
    'Point4
    dblStrPts(9) = dblStrPts(6)
    dblStrPts(10) = CenterPos.y
    dblStrPts(11) = parExchangerDiameter / 2
    'Point5
    dblStrPts(12) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    dblStrPts(13) = CenterPos.y
    dblStrPts(14) = parExchangerDiameter / 2
    'Point6
    dblStrPts(15) = (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    dblStrPts(16) = CenterPos.y
    dblStrPts(17) = parFrontEndFlangeDia / 2
    'Point7
    dblStrPts(18) = (parFrontEndLength1 + parBundleFlangeTk)
    dblStrPts(19) = CenterPos.y
    dblStrPts(20) = parFrontEndFlangeDia / 2
    'Point8
    dblStrPts(21) = (parFrontEndLength1 + parBundleFlangeTk)
    dblStrPts(22) = CenterPos.y
    dblStrPts(23) = parBundleFlangeDia / 2
    'Point9
    dblStrPts(24) = (parFrontEndLength1)
    dblStrPts(25) = CenterPos.y
    dblStrPts(26) = parBundleFlangeDia / 2
    'Point10
    dblStrPts(27) = (parFrontEndLength1)
    dblStrPts(28) = CenterPos.y
    dblStrPts(29) = CenterPos.z
    
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblStrPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (parFrontEndLength1 + parExchangerLength / 2), _
                                    CenterPos.y, _
                                    CenterPos.z

    Set ObjExchangerBody = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set oLineStr = Nothing

'Expansion Joint
    'P7 and P8 are defined provide expansion Joint
    If (parExpansionJointPosition > 0 And parExpansionJointThickness > 0) Then
        Dim ObjExpanJt  As Object
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set parExpansionJointPosition, CenterPos.y, CenterPos.z
        endPoint.Set parExpansionJointPosition + parExpansionJointThickness, CenterPos.y, CenterPos.z

        Set ObjExpanJt = PlaceCylinder(m_OutputColl, startPoint, endPoint, parExpansionJointDiameter, True)

 ' Set the output
        m_OutputColl.AddOutput "Expansion_", ObjExpanJt

        Set ObjExpanJt = Nothing
        Set startPoint = Nothing
        Set endPoint = Nothing
    End If
    'P16  is defined then place Support
    
    If (parSupportLength > 0) Then
        Dim oSupLine As IngrGeom3D.LineString3d
        Dim ObjDefsur  As IngrGeom3D.Plane3d
        Dim dblSupPts(0 To 14) As Double
        Dim projVect As New AutoMath.DVector
        Dim projHt As Double, ActualProj As Double
        'P15 is defined Place Bottom Support
        If (parBottomSupportHeight > 0) Then
            'P11 is defined Place First Bottom Support
            If (parBotSupportCenFromPP > 0) Then
                'Point 1        x (P11-P13/2)
                dblSupPts(0) = (parBotSupportCenFromPP - parSupport1Thickness / 2)
                dblSupPts(1) = -(parSupportLength / 2)
                dblSupPts(2) = -parBottomSupportHeight
                'Point2         x (P11+P13/2)
                dblSupPts(3) = dblSupPts(0) + parSupport1Thickness
                dblSupPts(4) = dblSupPts(1)
                dblSupPts(5) = -parBottomSupportHeight
                 'Point3
                dblSupPts(6) = dblSupPts(3)
                dblSupPts(7) = (parSupportLength / 2)
                dblSupPts(8) = -parBottomSupportHeight
                 'Point4
                dblSupPts(9) = dblSupPts(0)
                dblSupPts(10) = (parSupportLength / 2)
                dblSupPts(11) = -parBottomSupportHeight
                'Point5
                dblSupPts(12) = dblSupPts(0)
                dblSupPts(13) = dblSupPts(1)
                dblSupPts(14) = dblSupPts(2)
                projHt = Sqr((parExchangerDiameter / 2) * (parExchangerDiameter / 2) - (parSupportLength / 2) * (parSupportLength / 2))
                projVect.Set 0, 0, 1
                ActualProj = parBottomSupportHeight - projHt

                Set oSupLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
                Set ObjDefsur = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblSupPts)
                Set ObjSupport1 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, True)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport1
                m_OutputColl.AddOutput "DefaultSur_", ObjDefsur
                Set ObjSupport1 = Nothing
                Set ObjDefsur = Nothing
            End If              'P11 is Defined
            'P11 and P12 are defined Place 2nd Bottom Support
            If (parBotSupportCenFromPP > 0) And (parBottomSupportCentoCen > 0) Then
                ' Insert your code for output 3(Support2) Bottom Support
                'Point 1
                dblSupPts(0) = (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2)
                dblSupPts(1) = -(parSupportLength / 2)
                dblSupPts(2) = -parBottomSupportHeight
                'Point2
                dblSupPts(3) = dblSupPts(0) + parSupport2Thickness
                dblSupPts(4) = dblSupPts(1)
                dblSupPts(5) = -parBottomSupportHeight
                'Point3
                dblSupPts(6) = dblSupPts(3)
                dblSupPts(7) = (parSupportLength / 2)
                dblSupPts(8) = -parBottomSupportHeight
                'Point4
                dblSupPts(9) = dblSupPts(0)
                dblSupPts(10) = (parSupportLength / 2)
                dblSupPts(11) = -parBottomSupportHeight
                'Point5
                dblSupPts(12) = dblSupPts(0)
                dblSupPts(13) = dblSupPts(1)
                dblSupPts(14) = dblSupPts(2)
                projHt = Sqr((parExchangerDiameter / 2) * (parExchangerDiameter / 2) - (parSupportLength / 2) * (parSupportLength / 2))
                projVect.Set 0, 0, 1
                ActualProj = parBottomSupportHeight - projHt

                Set oSupLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
                Set ObjDefsur = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblSupPts)
                Set ObjSupport2 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, True)
            
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport2
                m_OutputColl.AddOutput "DefaultSur_", ObjDefsur
                Set ObjSupport2 = Nothing
                Set ObjDefsur = Nothing
            End If          'P11 and P12 are defined
        End If          'P15 is Defined
        ' Insert your code for output 4(Support3) Top Support
        'P19 is Defined Provide Top Support
        If (parTopSupportHeight > 0) Then
            'if P17 is defined then provide first Top Support
            If parTopSupportCenFromPP > 0 Then
                'Point 1
                dblSupPts(0) = (parTopSupportCenFromPP - parSupport1Thickness / 2)
                dblSupPts(1) = -(parSupportLength / 2)
                dblSupPts(2) = parTopSupportHeight
                'Point2
                dblSupPts(3) = dblSupPts(0) + parSupport1Thickness
                dblSupPts(4) = dblSupPts(1)
                dblSupPts(5) = parTopSupportHeight
                'Point3
                dblSupPts(6) = dblSupPts(3)
                dblSupPts(7) = (parSupportLength / 2)
                dblSupPts(8) = parTopSupportHeight
                'Point4
                dblSupPts(9) = dblSupPts(0)
                dblSupPts(10) = (parSupportLength / 2)
                dblSupPts(11) = parTopSupportHeight
                'Point5
                dblSupPts(12) = dblSupPts(0)
                dblSupPts(13) = dblSupPts(1)
                dblSupPts(14) = dblSupPts(2)
                projHt = Sqr((parExchangerDiameter / 2) * (parExchangerDiameter / 2) - (parSupportLength / 2) * (parSupportLength / 2))
                projVect.Set 0, 0, -1
                ActualProj = parTopSupportHeight - projHt

                Set oSupLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
                Set ObjSupport3 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, True)
            
            ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport3
                Set ObjSupport3 = Nothing
            End If
            'P17 and P18 are defined Provide second Top Support
            If (parTopSupportCenFromPP > 0) And (parTopSupportCentoCen > 0) Then
                ' Insert your code for output 5(Support4)
                'Point 1
                dblSupPts(0) = (parTopSupportCenFromPP + parTopSupportCentoCen - parSupport2Thickness / 2)
                dblSupPts(1) = -(parSupportLength / 2)
                dblSupPts(2) = parTopSupportHeight
                'Point2
                dblSupPts(3) = dblSupPts(0) + parSupport2Thickness
                dblSupPts(4) = dblSupPts(1)
                dblSupPts(5) = parTopSupportHeight
                'Point3
                dblSupPts(6) = dblSupPts(3)
                dblSupPts(7) = (parSupportLength / 2)
                dblSupPts(8) = parTopSupportHeight
                'Point4
                dblSupPts(9) = dblSupPts(0)
                dblSupPts(10) = (parSupportLength / 2)
                dblSupPts(11) = parTopSupportHeight
                'Point5
                dblSupPts(12) = dblSupPts(0)
                dblSupPts(13) = dblSupPts(1)
                dblSupPts(14) = dblSupPts(2)
                projHt = Sqr((parExchangerDiameter / 2) * (parExchangerDiameter / 2) - (parSupportLength / 2) * (parSupportLength / 2))
                projVect.Set 0, 0, -1
                ActualProj = parTopSupportHeight - projHt

                Set oSupLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
                Set ObjSupport4 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, True)
            
            ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport4
                Set ObjSupport4 = Nothing
                Set oSupLine = Nothing
                Set projVect = Nothing
            End If                              'P17 and P18 are defined
        End If                                      'P19 is defined
    End If                                           'P16 not Defined
    
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
    dblEndPts(5) = parFrontEndFlangeDia / 2     'P30/2
    'Point 3
    dblEndPts(6) = (parFrontEndLength1 - parFrontEndFlangeTk1)
    dblEndPts(7) = CenterPos.y
    dblEndPts(8) = parFrontEndFlangeDia / 2   'P30/2
    'Point 4
    dblEndPts(9) = dblEndPts(6)
    dblEndPts(10) = dblEndPts(7)
    dblEndPts(11) = parExchangerDiameter / 2
    'Point 5
    dblEndPts(12) = -(parFrontEndLength2 - parFrontEndFlangeTk3)
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = parExchangerDiameter / 2
    'Point 6
    dblEndPts(15) = dblEndPts(12)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parFrontEndFlangeDia / 2
    'Point 7
    dblEndPts(18) = -parFrontEndLength2
    dblEndPts(19) = CenterPos.y
    dblEndPts(20) = parFrontEndFlangeDia / 2
    'Point 8
    dblEndPts(21) = -(parFrontEndLength2 + parFrontEndFlangeTk2)
    dblEndPts(22) = CenterPos.y
    dblEndPts(23) = parFrontEndFlangeDia / 2
    'Point 9
    dblEndPts(24) = dblEndPts(21)
    dblEndPts(25) = CenterPos.y
    dblEndPts(26) = CenterPos.z
    
    Set oEndLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set 0, 0, 0

    Set ObjFrontEndBody = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontEndBody
    Set ObjFrontEndBody = Nothing
    
 ' Insert your code for output 7(RearEndBody)
    'Point 1 Start Point P32+P1
    Dim stPoint As Double
    stPoint = parExchangerLength + parFrontEndLength1
    dblEndPts(0) = stPoint
    dblEndPts(1) = CenterPos.y
    dblEndPts(2) = CenterPos.z
    'Point 2
    dblEndPts(3) = dblEndPts(0)
    dblEndPts(4) = dblEndPts(1)
    dblEndPts(5) = parRearEndFlangeDia / 2     'P40/2
    'Point 3
    dblEndPts(6) = (stPoint + parRearEndFlangeTk1)                  'P41
    dblEndPts(7) = CenterPos.y
    dblEndPts(8) = parRearEndFlangeDia / 2   'P40/2
    'Point 4
    dblEndPts(9) = dblEndPts(6)
    dblEndPts(10) = CenterPos.y
    dblEndPts(11) = parExchangerDiameter / 2               '(P40-P2)/2
    'Point 5
    dblEndPts(12) = (stPoint + parRearEndLength - parRearEndFlangeTk3)       'P42-P44
    dblEndPts(13) = CenterPos.y
    dblEndPts(14) = parExchangerDiameter / 2
    'Point 6
    dblEndPts(15) = dblEndPts(12)
    dblEndPts(16) = CenterPos.y
    dblEndPts(17) = parRearEndFlangeDia / 2
    'Point 7
    dblEndPts(18) = (stPoint + parRearEndLength)
    dblEndPts(19) = CenterPos.y
    dblEndPts(20) = parRearEndFlangeDia / 2
    'Point 8
    dblEndPts(21) = (stPoint + parRearEndLength + parRearEndFlangeTk2)              'P43
    dblEndPts(22) = CenterPos.y
    dblEndPts(23) = parRearEndFlangeDia / 2
    'Point 9
    dblEndPts(24) = dblEndPts(21)
    dblEndPts(25) = dblEndPts(22)
    dblEndPts(26) = CenterPos.z

    Set oEndLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set (stPoint + parRearEndLength / 2), CenterPos.y, CenterPos.z

    Set ObjRearEndBody = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearEndBody
    Set ObjRearEndBody = Nothing
    Set revVector = Nothing
    Set revCenterPt = Nothing
    Set oEndLineStr = Nothing

    '==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.05)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "HoriShellTubeExchangerControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    Set geomFactory = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

