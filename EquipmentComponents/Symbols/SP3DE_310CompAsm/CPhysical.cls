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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:51 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Friday, Feb 27 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'   E310 Symbol in Page no 305 of pdf file.  This symbol has variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
PI = 4 * Atn(1)
'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjExchangerBody As Object
    Dim ObjFrontEndBody As Object
    Dim ObjRearEndBody As Object
    Dim ObjExpanJoint As Object
    
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
    Dim parNumberOfSupports As Long
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
 ' Insert your code for output 1(ExchangerBody)
    'Form a line string and Revolve it about X-Axis
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblStrPts(0 To 29) As Double
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition

    'Point 1
    dblStrPts(0) = 0
    dblStrPts(1) = 0
    dblStrPts(2) = -(parFrontEndLength1 + parExchangerLength)                     '-(P32+P1)
    'Point2
    dblStrPts(3) = parRearEndFlangeDia / 2
    dblStrPts(4) = 0
    dblStrPts(5) = dblStrPts(2)
    'Point3
    dblStrPts(6) = dblStrPts(3)
    dblStrPts(7) = 0
    dblStrPts(8) = -(parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2)
    'Point4
    dblStrPts(9) = parExchangerDiameter / 2
    dblStrPts(10) = 0
    dblStrPts(11) = dblStrPts(8)
    'Point5
    dblStrPts(12) = dblStrPts(9)
    dblStrPts(13) = 0
    dblStrPts(14) = -(parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1)
    'Point6
    dblStrPts(15) = parFrontEndFlangeDia / 2
    dblStrPts(16) = 0
    dblStrPts(17) = dblStrPts(14)
    'Point7
    dblStrPts(18) = parFrontEndFlangeDia / 2
    dblStrPts(19) = 0
    dblStrPts(20) = -(parFrontEndLength1 + parBundleFlangeTk)
    'Point8
    dblStrPts(21) = parBundleFlangeDia / 2
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
    
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblStrPts)
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, (dblStrPts(2) + dblStrPts(29)) / 2

    Set ObjExchangerBody = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    
 ' Insert your code for output 2(FrontEndBody)
    Dim oEndLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 26) As Double
    
    'Point 1
    dblEndPts(0) = 0
    dblEndPts(1) = 0
    dblEndPts(2) = -parFrontEndLength1
    'Point 2
    dblEndPts(3) = parFrontEndFlangeDia / 2
    dblEndPts(4) = 0
    dblEndPts(5) = -parFrontEndLength1
    'Point 3
    dblEndPts(6) = dblEndPts(3)
    dblEndPts(7) = 0
    dblEndPts(8) = -(parFrontEndLength1 - parFrontEndFlangeTk1)
    'Point 4
    dblEndPts(9) = parExchangerDiameter / 2
    dblEndPts(10) = 0
    dblEndPts(11) = dblEndPts(8)
    'Point 5
    dblEndPts(12) = dblEndPts(9)
    dblEndPts(13) = 0
    dblEndPts(14) = parFrontEndLength2 - parFrontEndFlangeTk3
    'Point 6
    dblEndPts(15) = parFrontEndFlangeDia / 2
    dblEndPts(16) = 0
    dblEndPts(17) = dblEndPts(14)
    'Point 7
    dblEndPts(18) = parFrontEndFlangeDia / 2
    dblEndPts(19) = 0
    dblEndPts(20) = parFrontEndLength2
    'Point 8
    dblEndPts(21) = parFrontEndFlangeDia / 2
    dblEndPts(22) = 0
    dblEndPts(23) = parFrontEndLength2 + parFrontEndFlangeTk2
    'Point 9
    dblEndPts(24) = 0
    dblEndPts(25) = 0
    dblEndPts(26) = dblEndPts(23)
    
    Set oEndLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dblEndPts)
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, (dblEndPts(2) + dblEndPts(26)) / 2

    Set ObjFrontEndBody = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontEndBody
    Set ObjFrontEndBody = Nothing
    
 ' Insert your code for output 3(RearEndBody)
    'Point 1
    dblEndPts(0) = 0
    dblEndPts(1) = 0
    dblEndPts(2) = -(parFrontEndLength1 + parExchangerLength)
    'Point 2
    dblEndPts(3) = parRearEndFlangeDia / 2
    dblEndPts(4) = 0
    dblEndPts(5) = dblEndPts(2)
    'Point 3
    dblEndPts(6) = dblEndPts(3)
    dblEndPts(7) = 0
    dblEndPts(8) = -(parFrontEndLength1 + parExchangerLength + parRearEndFlangeTk1)
    'Point 4
    dblEndPts(9) = parExchangerDiameter / 2
    dblEndPts(10) = 0
    dblEndPts(11) = dblEndPts(8)
    'Point 5
    dblEndPts(12) = dblEndPts(9)
    dblEndPts(13) = 0
    dblEndPts(14) = -(parFrontEndLength1 + parExchangerLength + parRearEndLength - parRearEndFlangeTk3)
    'Point 6
    dblEndPts(15) = parRearEndFlangeDia / 2
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

    Set ObjRearEndBody = PlaceRevolution(m_OutputColl, oEndLineStr, revVector, revCenterPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearEndBody
    Set ObjRearEndBody = Nothing
    Set oEndLineStr = Nothing
    
 ' Insert your code for output 4(Extension joint)
    'P7 and P8 are defined provide Extension Joint
    If CmpDblGreaterthan(parExtensionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExtensionJointThickness, LINEAR_TOLERANCE) Then
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set 0, 0, -parExtensionJointPosition
        endPoint.Set 0, 0, -parExtensionJointPosition - parExtensionJointThickness

        Set ObjExpanJoint = PlaceCylinder(m_OutputColl, startPoint, endPoint, parExtensionJointDiameter, True)
    
    ' Set the output
        m_OutputColl.AddOutput "ExtensionJt_", ObjExpanJoint
        Set ObjExpanJoint = Nothing
        Set startPoint = Nothing
        Set endPoint = Nothing
    End If
    
'Insert your code for Support and Default Surface
    'For Support create Line Strings to represent top and bottom of Support and Use Ruled surfaces
    'to generate the final Shape.

    If (parNumberOfSupports > 0) Then
        If (CmpDblGreaterthan(parSupportDepth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportTopWidth, LINEAR_TOLERANCE)) Then
            Dim IncludedAng As Double
            Dim yAxistoSupport As Double
            Dim iCount As Double
            Dim ObjSupport As Object
            Dim ObjSurface As Object
            Dim ProVect As New AutoMath.DVector
            Dim SupPoints(0 To 14) As Double
            Dim SurPoints(0 To 14) As Double
            IncludedAng = (PI * 2) / parNumberOfSupports
            yAxistoSupport = parSupportAngularLocation
    
            For iCount = 1 To parNumberOfSupports
                'Point 1
                SupPoints(0) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) + Cos(yAxistoSupport) * parSupportThickness / 2
                SupPoints(1) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) - Sin(yAxistoSupport) * parSupportThickness / 2
                SupPoints(2) = -parSupportFromOrigin
                'Point 2
                SupPoints(3) = SupPoints(0) + Sin(yAxistoSupport) * (parSupportTopWidth - parExchangerDiameter) / 2
                SupPoints(4) = SupPoints(1) + Cos(yAxistoSupport) * (parSupportTopWidth - parExchangerDiameter) / 2
                SupPoints(5) = -parSupportFromOrigin
                'Point 4
                SupPoints(9) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) + Cos(yAxistoSupport) * parSupportThickness / 2
                SupPoints(10) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) - Sin(yAxistoSupport) * parSupportThickness / 2
                SupPoints(11) = -parSupportFromOrigin - parSupportDepth
                'Point3
                SupPoints(6) = SupPoints(9) + Sin(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SupPoints(7) = SupPoints(10) + Cos(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SupPoints(8) = -parSupportFromOrigin - parSupportDepth
                'Point5
                SupPoints(12) = SupPoints(0)
                SupPoints(13) = SupPoints(1)
                SupPoints(14) = -parSupportFromOrigin
                
                ProVect.Set -Cos(yAxistoSupport), Sin(yAxistoSupport), 0
                
                Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, SupPoints)
                Set ObjSupport = PlaceProjection(m_OutputColl, oLineStr, ProVect, parSupportThickness, True)
                 
                m_OutputColl.AddOutput "Support_", ObjSupport
                Set ObjSupport = Nothing
                 
                'Bottom Support Bottom Line String
                'Point 1
                SurPoints(0) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) + Cos(yAxistoSupport) * parSupportThickness / 2
                SurPoints(1) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) - Sin(yAxistoSupport) * parSupportThickness / 2
                SurPoints(2) = -parSupportFromOrigin - parSupportDepth
                'Point 2
                SurPoints(3) = SurPoints(0) + Sin(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(4) = SurPoints(1) + Cos(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(5) = -parSupportFromOrigin - parSupportDepth
                'Point 4
                SurPoints(9) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) - Cos(yAxistoSupport) * parSupportThickness / 2
                SurPoints(10) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) + Sin(yAxistoSupport) * parSupportThickness / 2
                SurPoints(11) = -parSupportFromOrigin - parSupportDepth
                'Point3
                SurPoints(6) = SurPoints(9) + Sin(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(7) = SurPoints(10) + Cos(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(8) = -parSupportFromOrigin - parSupportDepth
                'Point5
                SurPoints(12) = SurPoints(0)
                SurPoints(13) = SurPoints(1)
                SurPoints(14) = -parSupportFromOrigin - parSupportDepth
                
                Set ObjSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, SurPoints)
                
                'Set the Output
                m_OutputColl.AddOutput "Surface_", ObjSurface
                 
                'Increment the Angle
                yAxistoSupport = yAxistoSupport + IncludedAng
                Set ObjSupport = Nothing
                Set ObjSurface = Nothing
                
            Next iCount
            Set ProVect = Nothing
        End If
    Else                            'Place Ring or Skirt Support
        Dim centerPoint As New AutoMath.DPosition
        Dim ObjDefPlane As IngrGeom3D.Plane3d
        Dim oCircle  As IngrGeom3D.Circle3d
        Dim oCompstr As New ComplexString3d
        Dim baseCenter As New AutoMath.DPosition
        Dim topCenter As New AutoMath.DPosition
        
        baseCenter.Set 0, 0, -parSupportFromOrigin - parSupportDepth
        topCenter.Set 0, 0, baseCenter.z + parSupportDepth
        centerPoint.Set 0, 0, -parSupportFromOrigin - parSupportDepth
        
        Set ObjSupport = PlaceCone(m_OutputColl, baseCenter, topCenter, _
                                        parSupportBottomWidth / 2, parSupportTopWidth / 2, True)
        
        m_OutputColl.AddOutput "Support_", ObjSupport
        
        Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, centerPoint.x, _
                                                    centerPoint.y, centerPoint.z, 0, 0, 1, parSupportBottomWidth / 2)
                                                    
        Call oCompstr.AddCurve(oCircle, True)
        Set ObjDefPlane = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, centerPoint.x, centerPoint.y, _
                                                centerPoint.z, 0, 0, 1)
        Call ObjDefPlane.AddBoundary(oCompstr)
        
        m_OutputColl.AddOutput "Surface_", ObjDefPlane
        
        Dim ObjTmpcurves As IJDObject
        Set ObjTmpcurves = oCompstr
        ObjTmpcurves.Remove
        
        Set ObjDefPlane = Nothing
        Set centerPoint = Nothing
        Set oCircle = Nothing
        Set oCompstr = Nothing
        Set ObjSupport = Nothing
    
    End If
    
  
    Exit Sub
    
ErrorLabel:
      Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub
