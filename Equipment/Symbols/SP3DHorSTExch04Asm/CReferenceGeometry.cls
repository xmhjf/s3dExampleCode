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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:20 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Apr 25, 2003
'   Description:
'   This is a Shell and Tube Exchanger(E-182).
'   In this Symbol there are total 27 outputs.25 outputs are physical ,out of which 7 are pipe Nozzles.
'   There are 2 Reference Geometry objects which are Control-Point at origin and a reference plane at support bottom.
'   This is a Dow symbol Equipment
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim parVesselDiameter As Double
        Dim parVesselTantoTan As Double
        Dim parVesselLength As Double
        Dim parVesselStartPoint As Double
        Dim parFaceDiameter As Double
        Dim parFaceThickness As Double
        Dim parSupportWidth As Double
        Dim parSupportLength As Double
        Dim parSupportHeight  As Double
        Dim parFirstSupportLocation  As Double
        Dim parSecondSupportLocation As Double
        Dim parBasePlateThickness As Double
        Dim parBackingPlateThickness As Double
        Dim parBackingPlateWidth As Double
        Dim parStiffnerThickness As Double
        Dim parStiffnerThickness1 As Double
        Dim parNozzlePosition As Double
        Dim parNozzlePosition1 As Double
        Dim parNozzlePosition2 As Double
        Dim parNozzlePosition3 As Double
        Dim parNozzlePosition4 As Double
        Dim parNozzlePosition5 As Double
        Dim parNozzlePosition6 As Double
        Dim parRadialPosition As Double
        Dim parRadialPosition1 As Double
        Dim parRadialPosition2 As Double
        Dim parRadialPosition3 As Double
        Dim parRadialPosition4 As Double
        Dim parRadialPosition5 As Double
        Dim parRadialPosition6 As Double
        Dim parNozzleOrientation As Double
        Dim parNozzleOrientation1 As Double
        Dim parNozzleOrientation2 As Double
        Dim parNozzleOrientation3 As Double
        Dim parNozzleOrientation4 As Double
        Dim parNozzleOrientation5 As Double
        Dim parNozzleOrientation6 As Double

    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parVesselLength = arrayOfInputs(4)
    parVesselStartPoint = arrayOfInputs(5)
    parFaceDiameter = arrayOfInputs(6)
    parFaceThickness = arrayOfInputs(7)
    parSupportWidth = arrayOfInputs(8)
    parSupportLength = arrayOfInputs(9)
    parSupportHeight = arrayOfInputs(10)
    parFirstSupportLocation = arrayOfInputs(11)
    parSecondSupportLocation = arrayOfInputs(12)
    parBasePlateThickness = arrayOfInputs(13)
    parBackingPlateThickness = arrayOfInputs(14)
    parBackingPlateWidth = arrayOfInputs(15)
    parStiffnerThickness = arrayOfInputs(16)
    parStiffnerThickness1 = arrayOfInputs(17)

    iOutput = 0
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

' Insert your code for output 26( Point at origin )

    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    '==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, _
                                                                                                                 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment

    iOutput = iOutput + 1
    m_OutputColl.AddOutput "ControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

'   Create DefaultSurface - for output 27 (a square shape is chosen)
    Dim SideOfSquare As Double
    Dim HalfOfSide As Double
    Dim lineStrPts() As Double

    SideOfSquare = 0.1  'Surfac is considered to be 10 cm x 10 cm square
    HalfOfSide = SideOfSquare / 2

    ReDim lineStrPts(0 To 11) As Double
    lineStrPts(0) = parSecondSupportLocation - HalfOfSide
    lineStrPts(1) = HalfOfSide
    lineStrPts(2) = -parSupportHeight

    lineStrPts(3) = lineStrPts(0)
    lineStrPts(4) = -HalfOfSide
    lineStrPts(5) = lineStrPts(2)

    lineStrPts(6) = parSecondSupportLocation + HalfOfSide
    lineStrPts(7) = -HalfOfSide
    lineStrPts(8) = lineStrPts(2)

    lineStrPts(9) = lineStrPts(6)
    lineStrPts(10) = HalfOfSide
    lineStrPts(11) = lineStrPts(2)

    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPts)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing

    Set oGeomFactory = Nothing
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

