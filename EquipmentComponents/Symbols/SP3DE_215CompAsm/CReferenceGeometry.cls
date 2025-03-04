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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:10 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         ACM
'   Creation Date:  Tuesday, Feb 17 2004
'   Description:
'       This is a Simple Vertical Cylider Equipment component with Legs. This is developed based on PDS E-215.
'       This is a variable out put symbol. Number of Support legs is a variable.
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
    Dim iOutput     As Double

    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parVesselStartPoint  As Double
    Dim parInsulationThickness As Double
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)                      'P2
    parVesselDiameter = arrayOfInputs(3)                    'P3
    parSupportAngularLocation = arrayOfInputs(4)       'P4
    parNumberOfSupports = arrayOfInputs(5)      'P5
    parSupportHeight = arrayOfInputs(6)             'P6
    parSupportLength = arrayOfInputs(7)            'P7
    parSupportThickness = arrayOfInputs(8)      'P8
    parVesselStartPoint = arrayOfInputs(9)         'P9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
   
' Insert your code for output 7(Default Surface /plane at Control Point of vessel )
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "SimVerCylEqpLegCompControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    
' Insert your code for output 8(Default Surface /plane at Control Point of vessel )
    ' create non-persistent circle to use for creating default surface ---
    
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                0, 0, 0, _
                                0, 0, -1, _
                                0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    m_OutputColl.AddOutput "DefaultSurface", ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
    Set objCircle = Nothing
    Set m_oGBSFactory = Nothing
    Set oGeomFactory = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
Private Sub Class_Initialize()
''
End Sub

Private Sub Class_Terminate()
    ''
End Sub
