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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:34 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Apr 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is vessel skirt based on "Tower T-181" symbol
' of The DOW Chemical company.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
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
    Dim parVesselDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parSkirtHeight As Double
    Dim parSupportHeight As Double
    Dim parSupportDiameter As Double
    Dim parBasePlateThickness As Double
    Dim parBasePlateDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parSkirtBottomDiameter = arrayOfInputs(3)
    parSkirtHeight = arrayOfInputs(4)
    parSupportHeight = arrayOfInputs(5)
    parSupportDiameter = arrayOfInputs(6)
    parBasePlateThickness = arrayOfInputs(7)
    parBasePlateDiameter = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)

    iOutput = 0

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

'Place a Point3D at skirt top
    Dim ObjPoint As IngrGeom3D.Point3d
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double

    xValue = 0
    yValue = 0
    zValue = parSkirtHeight + parBasePlateThickness
    Set ObjPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, xValue, yValue, zValue)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPoint
    Set ObjPoint = Nothing

'Place a vertical Line
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    Dim ObjVertLine As IngrGeom3D.Line3d
    startPoint(0) = 0
    startPoint(1) = 0
    startPoint(2) = 0
    endPoint(0) = 0
    endPoint(1) = 0
    endPoint(2) = parSkirtBottomDiameter / 4
    Set ObjVertLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                            startPoint(0), startPoint(1), startPoint(2), _
                            endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertLine
    Set ObjVertLine = Nothing

'Place Top Default Surface
    Dim ObjTopDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double

    normalX = 0
    normalY = 0
    normalZ = 1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                xValue, yValue, zValue, _
                                                normalX, normalY, normalZ, _
                                                0.1 * parVesselDiameter / 2)
    Set ObjTopDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopDefaultSurface
    Set ObjTopDefaultSurface = Nothing
    Set oCircle = Nothing

'Place Bottom Default Surface
    Dim ObjBotDefaultSurface As IngrGeom3D.Plane3d

    xValue = 0
    yValue = 0
    zValue = 0

    normalX = 0
    normalY = 0
    normalZ = 1

    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                xValue, yValue, zValue, _
                                                normalX, normalY, normalZ, _
                                                0.1 * parVesselDiameter / 2)
    Set ObjBotDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotDefaultSurface
    Set ObjBotDefaultSurface = Nothing
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

