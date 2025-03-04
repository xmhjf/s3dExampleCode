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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:03 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This is Complex Vertical Cylindrical Equipment Skirt Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E205 Symbol in Page no 286.
'   Symbol is created using the following Outputs:
'   i)  4 standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: Vessel uses 'PlaceRevolution'
'       c) Two ReferenceGeometry aspect outputs: a Default Surface and a Control point
'   ii) Variable Outputs:
'        a) Support
'        b) Surface for the support and
'        c) Intermediate dome for shell section 3
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

    Dim iOutput     As Double

    Dim parVesselHeight1 As Double
    Dim parVesselDiameter1 As Double
    Dim parSupportHeight1 As Double
    Dim parVesselHeight2 As Double
    Dim parVesselDiameter2 As Double
    Dim parSupportHeight2 As Double
    Dim parVesselHeight3 As Double
    Dim parVesselDiameter3 As Double
    Dim parSupportHeight3 As Double
    Dim parVesselHeight4 As Double
    Dim parVesselDiameter4 As Double
    Dim parSkirtHeight As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parVesselReferencePoint As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight1 = arrayOfInputs(2)   'P2
    parVesselDiameter1 = arrayOfInputs(3) 'P3
    parSupportHeight1 = arrayOfInputs(4)  'P4
    parVesselHeight2 = arrayOfInputs(5)   'P5
    parVesselDiameter2 = arrayOfInputs(6) 'P6
    parSupportHeight2 = arrayOfInputs(7)  'P7
    parVesselHeight3 = arrayOfInputs(8)   'P8
    parVesselDiameter3 = arrayOfInputs(9) 'P9
    parSupportHeight3 = arrayOfInputs(10) 'P10
    parVesselHeight4 = arrayOfInputs(11)  'P11
    parVesselDiameter4 = arrayOfInputs(12)      'P12
    parSkirtHeight = arrayOfInputs(13)          'P13
    parSkirtTopDiameter = arrayOfInputs(14)     'P14
    parSkirtBottomDiameter = arrayOfInputs(15)  'P15
    parVesselReferencePoint = arrayOfInputs(16) 'P16
    parInsulationThickness = arrayOfInputs(17)

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    iOutput = 0
    
'' Origin is at DP/PP 1 at the base of Skirt.

' Insert your code for output 3 (DefaultSurface)
    Dim ObjDefSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oComplexStr As New IngrGeom3D.ComplexString3d
    Dim normalVect As New AutoMath.DVector
    
    Dim baseCenter As AutoMath.DPosition
    Set baseCenter = New AutoMath.DPosition
    
    baseCenter.Set 0, 0, 0
    normalVect.Set 0, 0, -1
    
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    baseCenter.x, baseCenter.y, baseCenter.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    0.1 * parSkirtBottomDiameter / 2)
    Call oComplexStr.AddCurve(oCircle, False)
    Set ObjDefSurface = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z)
    Call ObjDefSurface.AddBoundary(oComplexStr)

'   Set  the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing

    Dim ObjTmpcurves As IJDObject
    Set ObjTmpcurves = oComplexStr
    ObjTmpcurves.Remove
    Set oComplexStr = Nothing

    Set ObjTmpcurves = oCircle
    ObjTmpcurves.Remove
    Set ObjTmpcurves = Nothing
    Set oCircle = Nothing

    Set normalVect = Nothing
    Set baseCenter = Nothing
    Set oGeomFactory = Nothing

' Insert your code for output (Point at origin)
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                  0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub


