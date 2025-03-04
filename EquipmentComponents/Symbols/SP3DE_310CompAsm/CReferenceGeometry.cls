VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferenceGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:54 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
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
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



Option Explicit
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
   
    Dim iOutput     As Double
    
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
            
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
  ''Insert your code for output 8(DefaultSurface)    ''Place Default surface at Origin
    ' create non-persistent circle to use for creating default surface ---
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                0, 0, 0, _
                                0, 0, -1, _
                                0.1 * parExchangerDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    m_OutputColl.AddOutput "DefaultSurface", ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    
    Set geomFactory = Nothing

' Insert your code for output 9(Default Surface /plane at Control Point of vessel )
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "VertShellTubeExchangerControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
   Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
