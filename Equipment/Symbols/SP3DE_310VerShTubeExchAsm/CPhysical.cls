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
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, Mar 5 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Five Outputs:
'      In which Default Surface is created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt           As PartFacelets.IJDPart
    Dim ObjExchangerBody    As Object
    Dim ObjExpanJoint       As Object
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
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
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
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'Origin is at DP/PP1
    Dim geomFactory   As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for DefaultSurface    'Place Default surface at Origin
    'Create non-persistent circle to use for creating default surface ---
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                0, 0, 0, _
                                0, 0, -1, _
                                0.1 * parExchangerDiameter / 2)

    'Create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    
    'Insert your code for ExchangerBody
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
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set revVector = Nothing
    Set revCenterPt = Nothing
    
    'Insert your code for Expansion joint
    'P7 and P8 are defined provide Expansion Joint
    If CmpDblGreaterthan(parExpansionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExpansionJointThickness, LINEAR_TOLERANCE) Then
        Dim startPoint As New AutoMath.DPosition
        Dim endPoint As New AutoMath.DPosition
        
        startPoint.Set 0, 0, -parExpansionJointPosition
        endPoint.Set 0, 0, -parExpansionJointPosition - parExpansionJointThickness

        Set ObjExpanJoint = PlaceCylinder(m_OutputColl, startPoint, endPoint, parExpansionJointDiameter, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExpanJoint
        Set ObjExpanJoint = Nothing
        Set startPoint = Nothing
        Set endPoint = Nothing
    End If
    
    'Insert your code for Support(s)
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
                SupPoints(9) = SupPoints(0) + Sin(yAxistoSupport) * (parSupportTopWidth - parExchangerDiameter) / 2
                SupPoints(10) = SupPoints(1) + Cos(yAxistoSupport) * (parSupportTopWidth - parExchangerDiameter) / 2
                SupPoints(11) = -parSupportFromOrigin
                'Point 4
                SupPoints(3) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) + Cos(yAxistoSupport) * parSupportThickness / 2
                SupPoints(4) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) - Sin(yAxistoSupport) * parSupportThickness / 2
                SupPoints(5) = -parSupportFromOrigin - parSupportDepth
                'Point 3
                SupPoints(6) = SupPoints(3) + Sin(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SupPoints(7) = SupPoints(4) + Cos(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
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
                Set oLineStr = Nothing
                
                'Create the surface at the bottom of each support such that it's surface vector points down
                'Point 1
                SurPoints(0) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) + Cos(yAxistoSupport) * parSupportThickness / 2
                SurPoints(1) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) - Sin(yAxistoSupport) * parSupportThickness / 2
                SurPoints(2) = -parSupportFromOrigin - parSupportDepth
                'Point 2
                SurPoints(3) = Sin(yAxistoSupport) * (parExchangerDiameter / 2) - Cos(yAxistoSupport) * parSupportThickness / 2
                SurPoints(4) = Cos(yAxistoSupport) * (parExchangerDiameter / 2) + Sin(yAxistoSupport) * parSupportThickness / 2
                SurPoints(5) = -parSupportFromOrigin - parSupportDepth
                'Point 3
                SurPoints(6) = SurPoints(3) + Sin(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(7) = SurPoints(4) + Cos(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(8) = -parSupportFromOrigin - parSupportDepth
                'Point 4
                SurPoints(9) = SurPoints(0) + Sin(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(10) = SurPoints(1) + Cos(yAxistoSupport) * (parSupportBottomWidth - parExchangerDiameter) / 2
                SurPoints(11) = -parSupportFromOrigin - parSupportDepth
                
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
        Set baseCenter = Nothing
        Set topCenter = Nothing
        Set ObjSupport = Nothing
    End If
    Set geomFactory = Nothing
        
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
