VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CRefGeometry.cls
'   Author: VRK
'   Creation Date:  Monday, Feb 26 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Two Outputs:
'      In which One Default Surface, and ControlPoint are created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'******************************************************************************

Option Explicit
Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
   
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
   
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for DefaultSurface    ''Place Default surface at Origin
    'Create non-persistent circle to use for creating default surface ---
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, -1, _
                            0.1 * parExchangerDiameter / 2)
    
    'Create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                           (m_outputColl.ResourceManager, objCircle)
                           
    'Set the output
    m_outputColl.AddOutput "DefaultSurface", ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    
    'Insert your code for Control Point of vessel
    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_outputColl.ResourceManager, 0, 0, 0, 0.02, , , False)
    
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_outputColl.AddOutput "HorShTubeExControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    Set oGeomFactory = Nothing
   
   Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
