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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:07 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Friday, Apr 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'The Symbol has three aspects viz., Physical, Insulation and Maintenance.Physical Aspect consist of
'1.Exchanger body, Front End of Exchanger,Rear end of Exchanger, Extension Joint(variable), Supports(variable)
'and four Nozzles at bottom of Rear End of Exchanger, Radial Nozzle at the Exchanger body, along X-axis
'in Rear Exchanger End, along Y-Axis in Rear Exchanger End.
'Insulation Aspect consist of Insulation for Exchanger, Rear end,front end,extension joint and all Nozzles.
'Maintenance Aspect consist of cylinder along z-axis for bundle Pulling
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   06.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
PI = 4 * Atn(1)
'''
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
    Dim parNumberOfSupport As Double
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
    Dim parInsulationThickness  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)
    parExchangerDiameter = arrayOfInputs(3)
    parBundleFlangeTk = arrayOfInputs(4)
    parBundleFlangeDia = arrayOfInputs(5)
    parExchangerFlangeTk1 = arrayOfInputs(6)
    parExchangerFlangeTk2 = arrayOfInputs(7)
    parExtensionJointPosition = arrayOfInputs(8)
    parExtensionJointThickness = arrayOfInputs(9)
    parExtensionJointDiameter = arrayOfInputs(10)
    parBundlePullingLength = arrayOfInputs(11)
    parSupportAngularLocation = arrayOfInputs(12)
    parNumberOfSupport = arrayOfInputs(13)
    parSupportDepth = arrayOfInputs(14)
    parSupportFromOrigin = arrayOfInputs(15)
    parSupportTopWidth = arrayOfInputs(16)
    parSupportBottomWidth = arrayOfInputs(17)
    parSupportThickness = arrayOfInputs(18)
    parFrontEndFlangeDia = arrayOfInputs(19)
    parFrontEndFlangeTk1 = arrayOfInputs(20)
    parFrontEndLength1 = arrayOfInputs(21)
    parFrontEndLength2 = arrayOfInputs(22)
    parFrontEndFlangeTk2 = arrayOfInputs(23)
    parFrontEndFlangeTk3 = arrayOfInputs(24)
    parRearEndFlangeDia = arrayOfInputs(25)
    parRearEndFlangeTk1 = arrayOfInputs(26)
    parRearEndLength = arrayOfInputs(27)
    parRearEndFlangeTk2 = arrayOfInputs(28)
    parRearEndFlangeTk3 = arrayOfInputs(29)
    parInsulationThickness = arrayOfInputs(30)

    iOutput = 0

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
    
 ' Insert your code for output 6(Front End Body Ins)
    Dim oEndLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 26) As Double
    'Point 1
    dblEndPts(0) = 0
    dblEndPts(1) = 0
    dblEndPts(2) = -parFrontEndLength1
    'Point 2
    dblEndPts(3) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(4) = 0
    dblEndPts(5) = -parFrontEndLength1
    'Point 3
    dblEndPts(6) = dblEndPts(3)
    dblEndPts(7) = 0
    dblEndPts(8) = -(parFrontEndLength1 - parFrontEndFlangeTk1)
    'Point 4
    dblEndPts(9) = (parExchangerDiameter / 2) + parInsulationThickness
    dblEndPts(10) = 0
    dblEndPts(11) = dblEndPts(8)
    'Point 5
    dblEndPts(12) = dblEndPts(9)
    dblEndPts(13) = 0
    dblEndPts(14) = parFrontEndLength2 - parFrontEndFlangeTk3
    'Point 6
    dblEndPts(15) = (parFrontEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(16) = 0
    dblEndPts(17) = dblEndPts(14)
    'Point 7
    dblEndPts(18) = dblEndPts(15)
    dblEndPts(19) = 0
    dblEndPts(20) = parFrontEndLength2
    'Point 8
    dblEndPts(21) = dblEndPts(18)
    dblEndPts(22) = 0
    dblEndPts(23) = parFrontEndLength2 + parFrontEndFlangeTk2
    'Point 9
    dblEndPts(24) = 0
    dblEndPts(25) = 0
    dblEndPts(26) = dblEndPts(23)
    
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
    dblEndPts(8) = -(parFrontEndLength1 + parExchangerLength + parRearEndFlangeTk1)
    'Point 4
    dblEndPts(9) = (parExchangerDiameter / 2) + parInsulationThickness
    dblEndPts(10) = 0
    dblEndPts(11) = dblEndPts(8)
    'Point 5
    dblEndPts(12) = dblEndPts(9)
    dblEndPts(13) = 0
    dblEndPts(14) = -(parFrontEndLength1 + parExchangerLength + parRearEndLength - parRearEndFlangeTk3)
    'Point 6
    dblEndPts(15) = (parRearEndFlangeDia / 2) + parInsulationThickness
    dblEndPts(16) = 0
    dblEndPts(17) = dblEndPts(14)
    'Point 7
    dblEndPts(18) = dblEndPts(15)
    dblEndPts(19) = 0
    dblEndPts(20) = -(parFrontEndLength1 + parExchangerLength + parRearEndLength)
    'Point 8
    dblEndPts(21) = dblEndPts(15)
    dblEndPts(22) = 0
    dblEndPts(23) = dblEndPts(20) - parRearEndFlangeTk2
    'Point 9
    dblEndPts(24) = 0
    dblEndPts(25) = 0
    dblEndPts(26) = dblEndPts(23)
    
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
    If (parExtensionJointPosition > 0 And parExtensionJointThickness > 0) Then
        Dim ObjExpenJtIns As Object
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set 0, 0, -parExtensionJointPosition - parInsulationThickness
        endPoint.Set 0, 0, -parExtensionJointPosition - parExtensionJointThickness - parInsulationThickness

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
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
