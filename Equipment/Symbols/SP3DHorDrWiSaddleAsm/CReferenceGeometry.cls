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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:19 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Monday, May 26 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-162.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parVesselDiameter     As Double
    Dim parVesselTantoTan     As Double
    Dim parInsulationThickness As Double

    'Saddle inputs
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportWidth As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parBasePlateThickness As Double
    Dim parBackingPlateThickness As Double
    Dim parBackingPlateWidth As Double
    Dim parStiffnerThickness As Double
    Dim parStiffnerThickness1 As Double
    Dim parStiffnerSpacing As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '5 for Drum
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parFirstSupportLocation = arrayOfInputs(4)
    parSecondSupportLocation = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    '11 Saddle inputs
    parSupportLength = arrayOfInputs(7)
    parSupportHeight = arrayOfInputs(8)
    parSupportWidth = arrayOfInputs(9)
    parBasePlateThickness = arrayOfInputs(10)
    parBackingPlateThickness = arrayOfInputs(11)
    parBackingPlateWidth = arrayOfInputs(12)
    parStiffnerThickness = arrayOfInputs(13)
    parStiffnerThickness1 = arrayOfInputs(14)
    parStiffnerSpacing = arrayOfInputs(15)
    parXboltHole = arrayOfInputs(16)
    parYboltHole = arrayOfInputs(17)


    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim VesselOuterDiameter     As Double
    Dim shellThickness As Double

    Set geomFactory = New IngrGeom3D.GeometryFactory
    shellThickness = 0.0508         'Assumed value for Shell thickness
    VesselOuterDiameter = parVesselDiameter + 2 * shellThickness

    'Control Point
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                                                                                0, 0, 0, 0.02, , , False)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment

    'output
    iOutput = iOutput + 1

    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    'Default Surface
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d

    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double

    xValue = 0
    yValue = 0
    zValue = 0
    normalX = 1
    normalY = 0
    normalZ = 0
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                                xValue, yValue, zValue, _
                                                                                normalX, normalY, normalZ, _
                                                                                VesselOuterDiameter / 4)
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1

    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set oCircle = Nothing

    'Place a Horizontal Line along the axis of the Drum
    Dim ObjHoriLine As IngrGeom3D.Line3d
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double

    startPoint(0) = 0
    startPoint(1) = 0
    startPoint(2) = 0
    endPoint(0) = VesselOuterDiameter / 4
    endPoint(1) = 0
    endPoint(2) = 0
    Set ObjHoriLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, startPoint(0), startPoint(1), startPoint(2), _
                                                                                                    endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1

    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriLine
    Set ObjHoriLine = Nothing
    Set geomFactory = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub

