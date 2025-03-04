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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:37 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: VRK
'   Creation Date:  Monday, Feb 26 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has TwentyFive Outputs:
'      The Rectangular Foundation Ports are created at the bottom of plates.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'******************************************************************************

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     
     Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    'Define Front End types
    Const FrontEndType_A = 1
    Const FrontEndType_B = 2
    Const FrontEndType_C = 3
    Const FrontEndType_D = 4
    Const FrontEndType_N = 5
    Const FrontEndType_Q = 6
    
    'Define Rear End types
    Const RearEndType_L = 1
    Const RearEndType_M = 2
    Const RearEndType_N = 3
    Const RearEndType_P = 4
    Const RearEndType_Q = 5
    Const RearEndType_S = 6
    Const RearEndType_T = 7
    Const RearEndType_U = 8
    Const RearEndType_W1 = 9
    Const RearEndType_W2 = 10

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput As Double
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
    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    'Origin is at DP/PP1( Channel Inlet)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    If CmpDblEqual(parBundlePullingLength, LINEAR_TOLERANCE) Then parBundlePullingLength = parExchangerLength

    'Insert your code for Exchanger Body Cylinder
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim objExchangerCyl As Object
    oStartPoint.Set parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1, 0, 0
    oEndPoint.Set parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2, 0, 0
    Set objExchangerCyl = PlaceCylinder(m_outputColl, oStartPoint, oEndPoint, parExchangerDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objExchangerCyl
    Set objExchangerCyl = Nothing
                            
    'Insert your code for Exchanger Rear Flange
    Dim ObjExchangerRearFlange As Object
    oStartPoint.Set parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2, 0, 0
    oEndPoint.Set parFrontEndLength1 + parExchangerLength, 0, 0
    Set ObjExchangerRearFlange = PlaceCylinder(m_outputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerRearFlange
    Set ObjExchangerRearFlange = Nothing
 
    'Insert your code for Exchanger Front Flange1
    Dim ObjExchangerFrontFlange1 As Object
    oStartPoint.Set parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1, 0, 0
    oEndPoint.Set parFrontEndLength1 + parBundleFlangeTk, 0, 0
    Set ObjExchangerFrontFlange1 = PlaceCylinder(m_outputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerFrontFlange1
    Set ObjExchangerFrontFlange1 = Nothing
 
    'Insert your code for Exchanger Front Flange2
    Dim ObjExchangerFrontFlange2 As Object
    oStartPoint.Set parFrontEndLength1 + parBundleFlangeTk, 0, 0
    oEndPoint.Set parFrontEndLength1, 0, 0
    Set ObjExchangerFrontFlange2 = PlaceCylinder(m_outputColl, oStartPoint, oEndPoint, parBundleFlangeDia, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerFrontFlange2
    Set ObjExchangerFrontFlange2 = Nothing
  
    'Expansion Joint
    'P7 and P8 are defined provide expansion Joint
    If CmpDblGreaterthan(parExpansionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExpansionJointThickness, LINEAR_TOLERANCE) Then
        Dim ObjExpanJt  As Object
        oStartPoint.Set parExpansionJointPosition, CenterPos.y, CenterPos.z
        oEndPoint.Set parExpansionJointPosition + parExpansionJointThickness, CenterPos.y, CenterPos.z
        Set ObjExpanJt = PlaceCylinder(m_outputColl, oStartPoint, oEndPoint, parExpansionJointDiameter, True)
        'Set the output
        m_outputColl.AddOutput "Expansion_", ObjExpanJt
        Set ObjExpanJt = Nothing
    End If
    
    'P16  is defined then place Supports
    If CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        'P15 is defined Place Bottom Support
        If CmpDblGreaterthan(parBottomSupportHeight, LINEAR_TOLERANCE) Then
           'P11 is defined Place First Bottom Support
            If CmpDblGreaterthan(parBotSupportCenFromPP, LINEAR_TOLERANCE) Then
                Dim ObjSupport1 As Object
                oStartPoint.Set parBotSupportCenFromPP - parSupport1Thickness / 2, _
                                        -(parSupportLength / 2), _
                                        -parBottomSupportHeight
                oEndPoint.Set oStartPoint.x + parSupport1Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parBottomSupportHeight
                                        
                Set ObjSupport1 = PlaceBox(m_outputColl, oStartPoint, oEndPoint)
                'Set the output
                m_outputColl.AddOutput "Supports_", ObjSupport1
                Set ObjSupport1 = Nothing
                       
            End If              'P11 is Defined
                        
            'P12 is defined Place 2nd Bottom Support
            If CmpDblGreaterthan(parBottomSupportCentoCen, LINEAR_TOLERANCE) Then
                'Insert your code for output 3(Support2) Bottom Support
                Dim ObjSupport2 As Object
                oStartPoint.Set (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2), _
                                       -(parSupportLength / 2), _
                                       -parBottomSupportHeight
                oEndPoint.Set oStartPoint.x + parSupport2Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parBottomSupportHeight
                                        
                Set ObjSupport2 = PlaceBox(m_outputColl, oStartPoint, oEndPoint)
                'Set the output
                m_outputColl.AddOutput "Supports_", ObjSupport2
                Set ObjSupport2 = Nothing
                
            End If          '' P12 is defined
            
        End If          'P15 is Defined
        
        ' Insert your code for Support3 Top Support
        'P19 is Defined Provide Top Support
        If CmpDblGreaterthan(parTopSupportHeight, LINEAR_TOLERANCE) Then
            'If P17 is defined then provide first Top Support
            If CmpDblGreaterthan(parTopSupportCenFromPP, LINEAR_TOLERANCE) Then
                Dim ObjSupport3 As Object
                oStartPoint.Set (parTopSupportCenFromPP - parSupport1Thickness / 2), _
                                       -(parSupportLength / 2), _
                                       0
                oEndPoint.Set oStartPoint.x + parSupport1Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parTopSupportHeight
                                        
                Set ObjSupport3 = PlaceBox(m_outputColl, oStartPoint, oEndPoint)
                'Set the output
                m_outputColl.AddOutput "Supports_", ObjSupport3
                Set ObjSupport3 = Nothing
            End If
            'P18 is defined Provide second Top Support
            If CmpDblGreaterthan(parTopSupportCentoCen, LINEAR_TOLERANCE) Then
                ' Insert your code for Support4
                
                Dim ObjSupport4 As Object
                
                oStartPoint.Set (parTopSupportCenFromPP + parTopSupportCentoCen - parSupport2Thickness / 2), _
                                       -(parSupportLength / 2), _
                                       0
                oEndPoint.Set oStartPoint.x + parSupport2Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parTopSupportHeight
                                        
                Set ObjSupport4 = PlaceBox(m_outputColl, oStartPoint, oEndPoint)
                
                ' Set the output
                m_outputColl.AddOutput "Supports_", ObjSupport4
                Set ObjSupport4 = Nothing
                                
            End If                              'P18 is defined
        End If                                      'P19 is defined
    End If                                           'P16 is Defined
    
    Set oGeomFactory = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    Set CenterPos = Nothing
    
    'Create Points on the top of supports:
    Dim objPoint As Object
    Dim centerPoint As IJDPosition
    Set centerPoint = New DPosition
    'P16  is defined then place Supports
    If CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        'P19 is Defined Provide Top Support
            If CmpDblGreaterthan(parTopSupportHeight, LINEAR_TOLERANCE) Then
                'If P17 is defined then provide first Top Support
                If CmpDblGreaterthan(parTopSupportCenFromPP, LINEAR_TOLERANCE) Then
                    'Create Connect Point on TopSupport1
                    centerPoint.Set parTopSupportCenFromPP, 0, parTopSupportHeight
                    Set objPoint = CreatePoint(m_outputColl, centerPoint, Nothing)
                    m_outputColl.AddOutput "Point1_", objPoint
                    Set objPoint = Nothing
                End If
            'P18 is defined Provide second Top Support
            If CmpDblGreaterthan(parTopSupportCentoCen, LINEAR_TOLERANCE) Then
                'Create Connect Point on TopSupport2
                centerPoint.Set parTopSupportCenFromPP + parTopSupportCentoCen, 0, parTopSupportHeight
                Set objPoint = CreatePoint(m_outputColl, centerPoint, Nothing)
                m_outputColl.AddOutput "Point2_", objPoint
                Set objPoint = Nothing
                
            End If                              'P18 is defined
        End If                                      'P19 is defined
End If                                           'P16 is Defined
    
    'Adding Edges to all Supports
    Dim iCount As Integer
    Dim Support1Points(0 To 23) As Double
    Dim Support2Points(0 To 23) As Double
    Dim Support3Points(0 To 23) As Double
    Dim Support4Points(0 To 23) As Double
    Dim ObjEdgeColl As Collection
    Set ObjEdgeColl = New Collection
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    If CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
             'P15 is defined Place Bottom Support
        If CmpDblGreaterthan(parBottomSupportHeight, LINEAR_TOLERANCE) Then
             'P11 is defined Place First Bottom Support
            If CmpDblGreaterthan(parBotSupportCenFromPP, LINEAR_TOLERANCE) Then
            'Support1
            centerPoint.Set parBotSupportCenFromPP, 0, -parBottomSupportHeight
            
             For iCount = 0 To 12 Step 12
                Support1Points(0 + iCount) = centerPoint.x - parSupport1Thickness / 2
                Support1Points(1 + iCount) = centerPoint.y - parSupportLength / 2
                Support1Points(2 + iCount) = centerPoint.z
        
                Support1Points(3 + iCount) = centerPoint.x - parSupport1Thickness / 2
                Support1Points(4 + iCount) = centerPoint.y + parSupportLength / 2
                Support1Points(5 + iCount) = centerPoint.z
        
                Support1Points(6 + iCount) = centerPoint.x + parSupport1Thickness / 2
                Support1Points(7 + iCount) = centerPoint.y + parSupportLength / 2
                Support1Points(8 + iCount) = centerPoint.z
        
                Support1Points(9 + iCount) = centerPoint.x + parSupport1Thickness / 2
                Support1Points(10 + iCount) = centerPoint.y - parSupportLength / 2
                Support1Points(11 + iCount) = centerPoint.z
        
                centerPoint.z = centerPoint.z + parBottomSupportHeight
            Next iCount
        
            'Create Edge by two SupportPoints:
            'Edge1 Point1 to Point2
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(0), Support1Points(1), Support1Points(2), _
                                        Support1Points(3), Support1Points(4), Support1Points(5))
            'Edge2 Point1 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(0), Support1Points(1), Support1Points(2), _
                                        Support1Points(9), Support1Points(10), Support1Points(11))
            'Edge3 Point2 to Point3
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(3), Support1Points(4), Support1Points(5), _
                                        Support1Points(6), Support1Points(7), Support1Points(8))
            'Edge4 Point3 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(6), Support1Points(7), Support1Points(8), _
                                        Support1Points(9), Support1Points(10), Support1Points(11))
            'Edge5 Point5 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(12), Support1Points(13), Support1Points(14), _
                                        Support1Points(15), Support1Points(16), Support1Points(17))
            'Edge6 Point5 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(12), Support1Points(13), Support1Points(14), _
                                        Support1Points(21), Support1Points(22), Support1Points(23))
            'Edge7 Point6 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(15), Support1Points(16), Support1Points(17), _
                                        Support1Points(18), Support1Points(19), Support1Points(20))
            'Edge8 Point7 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(18), Support1Points(19), Support1Points(20), _
                                        Support1Points(21), Support1Points(22), Support1Points(23))
            'Edge9 Point1 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(0), Support1Points(1), Support1Points(2), _
                                        Support1Points(12), Support1Points(13), Support1Points(14))
            'Edge10 Point2 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(3), Support1Points(4), Support1Points(5), _
                                        Support1Points(15), Support1Points(16), Support1Points(17))
            'Edge11 Point3 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(6), Support1Points(7), Support1Points(8), _
                                        Support1Points(18), Support1Points(19), Support1Points(20))
            'Edge12 Point4 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support1Points(9), Support1Points(10), Support1Points(11), _
                                        Support1Points(21), Support1Points(22), Support1Points(23))
            'Set the Output
            For iCount = 1 To ObjEdgeColl.Count
                m_outputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
            Next iCount
            Set ObjEdgeColl = Nothing
        End If              'P11 is Defined
                        
    'P12 is defined Place 2nd Bottom Support
    If CmpDblGreaterthan(parBottomSupportCentoCen, LINEAR_TOLERANCE) Then
            'Support2
            centerPoint.Set parBotSupportCenFromPP + parBottomSupportCentoCen, 0, -parBottomSupportHeight
        
             For iCount = 0 To 12 Step 12
                Support2Points(0 + iCount) = centerPoint.x - parSupport2Thickness / 2
                Support2Points(1 + iCount) = centerPoint.y - parSupportLength / 2
                Support2Points(2 + iCount) = centerPoint.z
        
                Support2Points(3 + iCount) = centerPoint.x - parSupport2Thickness / 2
                Support2Points(4 + iCount) = centerPoint.y + parSupportLength / 2
                Support2Points(5 + iCount) = centerPoint.z
        
                Support2Points(6 + iCount) = centerPoint.x + parSupport2Thickness / 2
                Support2Points(7 + iCount) = centerPoint.y + parSupportLength / 2
                Support2Points(8 + iCount) = centerPoint.z
        
                Support2Points(9 + iCount) = centerPoint.x + parSupport2Thickness / 2
                Support2Points(10 + iCount) = centerPoint.y - parSupportLength / 2
                Support2Points(11 + iCount) = centerPoint.z
        
                centerPoint.z = centerPoint.z + parBottomSupportHeight
            Next iCount
        
            'Create Edge by two SupportPoints:
            Set ObjEdgeColl = New Collection
            'Edge1 Point1 to Point2
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(0), Support2Points(1), Support2Points(2), _
                                        Support2Points(3), Support2Points(4), Support2Points(5))
            'Edge2 Point1 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(0), Support2Points(1), Support2Points(2), _
                                        Support2Points(9), Support2Points(10), Support2Points(11))
            'Edge3 Point2 to Point3
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(3), Support2Points(4), Support2Points(5), _
                                        Support2Points(6), Support2Points(7), Support2Points(8))
            'Edge4 Point3 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(6), Support2Points(7), Support2Points(8), _
                                        Support2Points(9), Support2Points(10), Support2Points(11))
            'Edge5 Point5 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(12), Support2Points(13), Support2Points(14), _
                                        Support2Points(15), Support2Points(16), Support2Points(17))
            'Edge6 Point5 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(12), Support2Points(13), Support2Points(14), _
                                        Support2Points(21), Support2Points(22), Support2Points(23))
            'Edge7 Point6 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(15), Support2Points(16), Support2Points(17), _
                                        Support2Points(18), Support2Points(19), Support2Points(20))
            'Edge8 Point7 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(18), Support2Points(19), Support2Points(20), _
                                        Support2Points(21), Support2Points(22), Support2Points(23))
            'Edge9 Point1 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(0), Support2Points(1), Support2Points(2), _
                                        Support2Points(12), Support2Points(13), Support2Points(14))
            'Edge10 Point2 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(3), Support2Points(4), Support2Points(5), _
                                        Support2Points(15), Support2Points(16), Support2Points(17))
            'Edge11 Point3 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(6), Support2Points(7), Support2Points(8), _
                                        Support2Points(18), Support2Points(19), Support2Points(20))
            'Edge12 Point4 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support2Points(9), Support2Points(10), Support2Points(11), _
                                        Support2Points(21), Support2Points(22), Support2Points(23))
            'Set the Output
            For iCount = 1 To ObjEdgeColl.Count
                m_outputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
            Next iCount
            Set ObjEdgeColl = Nothing
          End If          '' P12 is defined
    End If          'P15 is Defined

    'P19 is Defined Provide Top Support
    If CmpDblGreaterthan(parTopSupportHeight, LINEAR_TOLERANCE) Then
         'If P17 is defined then provide first Top Support
        If CmpDblGreaterthan(parTopSupportCenFromPP, LINEAR_TOLERANCE) Then

            'Support3
            centerPoint.Set parTopSupportCenFromPP, 0, 0
        
             For iCount = 0 To 12 Step 12
                Support3Points(0 + iCount) = centerPoint.x - parSupport1Thickness / 2
                Support3Points(1 + iCount) = centerPoint.y - parSupportLength / 2
                Support3Points(2 + iCount) = centerPoint.z
        
                Support3Points(3 + iCount) = centerPoint.x - parSupport1Thickness / 2
                Support3Points(4 + iCount) = centerPoint.y + parSupportLength / 2
                Support3Points(5 + iCount) = centerPoint.z
        
                Support3Points(6 + iCount) = centerPoint.x + parSupport1Thickness / 2
                Support3Points(7 + iCount) = centerPoint.y + parSupportLength / 2
                Support3Points(8 + iCount) = centerPoint.z
        
                Support3Points(9 + iCount) = centerPoint.x + parSupport1Thickness / 2
                Support3Points(10 + iCount) = centerPoint.y - parSupportLength / 2
                Support3Points(11 + iCount) = centerPoint.z
        
                centerPoint.z = centerPoint.z + parTopSupportHeight
            Next iCount
        
            'Create Edge by two SupportPoints:
            Set ObjEdgeColl = New Collection
            'Edge1 Point1 to Point2
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(0), Support3Points(1), Support3Points(2), _
                                        Support3Points(3), Support3Points(4), Support3Points(5))
            'Edge2 Point1 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(0), Support3Points(1), Support3Points(2), _
                                        Support3Points(9), Support3Points(10), Support3Points(11))
            'Edge3 Point2 to Point3
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(3), Support3Points(4), Support3Points(5), _
                                        Support3Points(6), Support3Points(7), Support3Points(8))
            'Edge4 Point3 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(6), Support3Points(7), Support3Points(8), _
                                        Support3Points(9), Support3Points(10), Support3Points(11))
            'Edge5 Point5 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(12), Support3Points(13), Support3Points(14), _
                                        Support3Points(15), Support3Points(16), Support3Points(17))
            ''Edge6 Point5 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(12), Support3Points(13), Support3Points(14), _
                                        Support3Points(21), Support3Points(22), Support3Points(23))
            'Edge7 Point6 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(15), Support3Points(16), Support3Points(17), _
                                        Support3Points(18), Support3Points(19), Support3Points(20))
            'Edge8 Point7 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(18), Support3Points(19), Support3Points(20), _
                                        Support3Points(21), Support3Points(22), Support3Points(23))
            'Edge9 Point1 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(0), Support3Points(1), Support3Points(2), _
                                        Support3Points(12), Support3Points(13), Support3Points(14))
            'Edge10 Point2 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(3), Support3Points(4), Support3Points(5), _
                                        Support3Points(15), Support3Points(16), Support3Points(17))
            'Edge11 Point3 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(6), Support3Points(7), Support3Points(8), _
                                        Support3Points(18), Support3Points(19), Support3Points(20))
            'Edge12 Point4 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support3Points(9), Support3Points(10), Support3Points(11), _
                                        Support3Points(21), Support3Points(22), Support3Points(23))
            'Set the Output
            For iCount = 1 To ObjEdgeColl.Count
                m_outputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
            Next iCount
            Set ObjEdgeColl = Nothing
            
     End If
'P18 is defined Provide second Top Support
If CmpDblGreaterthan(parTopSupportCentoCen, LINEAR_TOLERANCE) Then

           'Support4
            centerPoint.Set parTopSupportCenFromPP + parTopSupportCentoCen, 0, 0
        
             For iCount = 0 To 12 Step 12
                Support4Points(0 + iCount) = centerPoint.x - parSupport2Thickness / 2
                Support4Points(1 + iCount) = centerPoint.y - parSupportLength / 2
                Support4Points(2 + iCount) = centerPoint.z
        
                Support4Points(3 + iCount) = centerPoint.x - parSupport2Thickness / 2
                Support4Points(4 + iCount) = centerPoint.y + parSupportLength / 2
                Support4Points(5 + iCount) = centerPoint.z
        
                Support4Points(6 + iCount) = centerPoint.x + parSupport2Thickness / 2
                Support4Points(7 + iCount) = centerPoint.y + parSupportLength / 2
                Support4Points(8 + iCount) = centerPoint.z
        
                Support4Points(9 + iCount) = centerPoint.x + parSupport2Thickness / 2
                Support4Points(10 + iCount) = centerPoint.y - parSupportLength / 2
                Support4Points(11 + iCount) = centerPoint.z
        
                centerPoint.z = centerPoint.z + parTopSupportHeight
            Next iCount
        
            'Create Edge by two SupportPoints:
            Set ObjEdgeColl = New Collection
            'Edge1 Point1 to Point2
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(0), Support4Points(1), Support4Points(2), _
                                        Support4Points(3), Support4Points(4), Support4Points(5))
            'Edge2 Point1 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(0), Support4Points(1), Support4Points(2), _
                                        Support4Points(9), Support4Points(10), Support4Points(11))
            'Edge3 Point2 to Point3
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(3), Support4Points(4), Support4Points(5), _
                                        Support4Points(6), Support4Points(7), Support4Points(8))
            'Edge4 Point3 to Point4
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(6), Support4Points(7), Support4Points(8), _
                                        Support4Points(9), Support4Points(10), Support4Points(11))
            'Edge5 Point5 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(12), Support4Points(13), Support4Points(14), _
                                        Support4Points(15), Support4Points(16), Support4Points(17))
            'Edge6 Point5 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(12), Support4Points(13), Support4Points(14), _
                                        Support4Points(21), Support4Points(22), Support4Points(23))
            'Edge7 Point6 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(15), Support4Points(16), Support4Points(17), _
                                        Support4Points(18), Support4Points(19), Support4Points(20))
            'Edge8 Point7 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(18), Support4Points(19), Support4Points(20), _
                                        Support4Points(21), Support4Points(22), Support4Points(23))
            'Edge9 Point1 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(0), Support4Points(1), Support4Points(2), _
                                        Support4Points(12), Support4Points(13), Support4Points(14))
            'Edge10 Point2 to Point6
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(3), Support4Points(4), Support4Points(5), _
                                        Support4Points(15), Support4Points(16), Support4Points(17))
            'Edge11 Point3 to Point7
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(6), Support4Points(7), Support4Points(8), _
                                        Support4Points(18), Support4Points(19), Support4Points(20))
            'Edge12 Point4 to Point8
            ObjEdgeColl.Add geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        Support4Points(9), Support4Points(10), Support4Points(11), _
                                        Support4Points(21), Support4Points(22), Support4Points(23))
            'Set the Output
            For iCount = 1 To ObjEdgeColl.Count
                m_outputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
            Next iCount
            Set ObjEdgeColl = Nothing
            Set centerPoint = Nothing
            End If                              'P18 is defined
    End If                                      'P19 is defined
End If                                           'P16 is Defined
 
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Coordinate system is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |       |
    '   |                  |       |
    '   |                  |       |
    '   -----> X           |       |
    '   Symbol CS          |       |
    '                      |       |
    '                      |   X   |
    '                      |   ^   |
    '                      |   |   |
    '                      |   |   |
    '                      |---|---|----> Y
    '                   Port CS

    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double

   'The origin of the port is taken to be at the corner point of the support base.
    dOrigin(0) = parBotSupportCenFromPP
    dOrigin(1) = 0
    dOrigin(2) = -parBottomSupportHeight
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_outputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parFP1XBoltHole
    holes(0, 2) = -parFP1YBoltHole
    holes(1, 1) = parFP1XBoltHole
    holes(1, 2) = -parFP1YBoltHole
    holes(2, 1) = parFP1XBoltHole
    holes(2, 2) = parFP1YBoltHole
    holes(3, 1) = -parFP1XBoltHole
    holes(3, 2) = parFP1YBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    m_outputColl.AddOutput "EqpFoundationPort1", ObjFoundationPort
    Set ObjFoundationPort = Nothing

    'The origin of the port is taken to be at the corner point of the support base.
    dOrigin(0) = parBotSupportCenFromPP + parBottomSupportCentoCen
    dOrigin(1) = 0
    dOrigin(2) = -parBottomSupportHeight

    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort2", _
                                                    False, m_outputColl.ResourceManager)

    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parFP2XBoltHole
    holes(0, 2) = -parFP2YBoltHole
    holes(1, 1) = parFP2XBoltHole
    holes(1, 2) = -parFP2YBoltHole
    holes(2, 1) = parFP2XBoltHole
    holes(2, 2) = parFP2YBoltHole
    holes(3, 1) = -parFP2XBoltHole
    holes(3, 2) = parFP2YBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    m_outputColl.AddOutput "EqpFoundationPort2", ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

