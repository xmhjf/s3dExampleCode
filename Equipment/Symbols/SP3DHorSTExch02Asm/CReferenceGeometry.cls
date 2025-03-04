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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:29 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Apr 25, 2003
'   Description:
'   This is a Shell and Tube Exchanger(E-162).
'   In this Symbol there are total 16 outputs.14 outputs are physical out of which 6 are pipe Nozzles.
'   There are 2 Reference Geometry objects which are Point at origin and a reference plane at support bottom.
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
    Dim parVesselLength As Double
    Dim parVesselTantoTan As Double
    Dim parVesselLength1 As Double
    Dim parVesselStartPoint As Double
    Dim parSupportLength As Double
    Dim parSupportWidth As Double
    Dim parSupportHeight  As Double
    Dim parFirstSupportLocation  As Double
    Dim parSecondSupportLocation As Double
    
    
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselLength = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)
    parVesselLength1 = arrayOfInputs(5)
    parVesselStartPoint = arrayOfInputs(6)
    parSupportLength = arrayOfInputs(7)
    parSupportWidth = arrayOfInputs(8)
    parSupportHeight = arrayOfInputs(9)
    parFirstSupportLocation = arrayOfInputs(10)
    parSecondSupportLocation = arrayOfInputs(11)
    
    
    iOutput = 0
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

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
    m_OutputColl.AddOutput "ControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
'   Create DefaultSurface - for output 16 (a square shape is chosen)
    Dim SideOfSquare As Double
    Dim HalfOfSide As Double
    Dim lineStrPts() As Double

    SideOfSquare = 0.4  'Surfac is considered to be 40 cm x 40 cm square
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

