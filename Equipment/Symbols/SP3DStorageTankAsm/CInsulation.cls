VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:39 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Feb 6 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This Symbol is a replicate of "Tank.TankServices" Symbol, except that the symbol is in line with
' Symbol created using wizard.
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------   -----                    ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   18.May.2005     kkk                 CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
'  08.SEP.2006      KKC               DI-95670  Replace names with initials in all revision history sheets and symbols
'   20.SEP.2006      KKC                TR-89336  Insulation Aspect is not shown correctly on the storage tank Elliptical ends.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
'''
        PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim ObjEndCap As IngrGeom3D.Revolution3d
    Dim Objcylinder As IngrGeom3D.Revolution3d
    
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportThickness As Double
    Dim parInsulationThickness As Double
    Dim dtolerance As Double
    Dim ellipseCenter As AutoMath.DPosition
    Dim ellipseNormal As AutoMath.DPosition
    Dim majorAxis As AutoMath.DPosition
    Dim daxesratio As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSupportLength = arrayOfInputs(4)
    parSupportHeight = arrayOfInputs(5)
    parSupportThickness = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    Set majorAxis = New AutoMath.DPosition
    Set ellipseNormal = New AutoMath.DPosition
    Set ellipseCenter = New AutoMath.DPosition
    
    iOutput = 0
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oCmplxString As IngrGeom3D.ComplexString3d
    Dim Linecurve As Object
    Dim ellipse As Object
    Dim ellipse1 As Object
    Dim planePoints(0 To 6) As Double
    Dim revVect As AutoMath.DVector
    Dim centerPoint As AutoMath.DPosition
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set revVect = New AutoMath.DVector
    Set centerPoint = New AutoMath.DPosition
    
    'Elliptical head minor radius is 0.5 times the Tank Radius
    
    ellipseCenter.Set 0, 0, parSupportHeight + parVesselDiameter * 0.5
    majorAxis.Set 0, 0, parVesselDiameter / 2 + parInsulationThickness
    ellipseNormal.Set 0, 1, 0
    daxesratio = (0.5 * 0.5 * parVesselDiameter + parInsulationThickness) _
                    / (0.5 * parVesselDiameter + parInsulationThickness)
    
    
    Set ellipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                ellipseCenter.x, ellipseCenter.y, ellipseCenter.z, _
                                ellipseNormal.x, ellipseNormal.y, ellipseNormal.z, _
                                majorAxis.x, majorAxis.y, majorAxis.z, _
                                daxesratio, 1.5 * PI, PI / 2)
                                
    revVect.Set 1, 0, 0
    centerPoint.Set 0, 0, parSupportHeight + parVesselDiameter * 0.5
    Set ObjEndCap = PlaceRevolution(m_OutputColl, ellipse, revVect, centerPoint, PI * 2, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEndCap
                              
    planePoints(0) = 0
    planePoints(1) = 0
    planePoints(2) = parSupportHeight + parVesselDiameter + parInsulationThickness
    
    planePoints(3) = parVesselLength
    planePoints(4) = 0
    planePoints(5) = parSupportHeight + parVesselDiameter + parInsulationThickness
    
    Set Linecurve = geomFactory.Lines3d.CreateBy2Points(Nothing, planePoints(0), planePoints(1), planePoints(2), _
                                                     planePoints(3), planePoints(4), planePoints(5))
    
    
    revVect.Set 1, 0, 0
    centerPoint.Set parVesselLength * 0.5, 0, parSupportHeight + parVesselDiameter * 0.5
    Set Objcylinder = PlaceRevolution(m_OutputColl, Linecurve, revVect, centerPoint, PI * 2, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), Objcylinder
    
    ellipseCenter.Set parVesselLength, 0, parSupportHeight + parVesselDiameter * 0.5
    majorAxis.Set 0, 0, parVesselDiameter / 2 + parInsulationThickness
    ellipseNormal.Set 0, -1, 0
    
    daxesratio = (0.5 * 0.5 * parVesselDiameter + parInsulationThickness) _
                    / (0.5 * parVesselDiameter + parInsulationThickness)
    Set ellipse1 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                ellipseCenter.x, ellipseCenter.y, ellipseCenter.z, _
                                ellipseNormal.x, ellipseNormal.y, ellipseNormal.z, _
                                majorAxis.x, majorAxis.y, majorAxis.z, _
                                daxesratio, 1.5 * PI, PI / 2)
    revVect.Set 1, 0, 0
    centerPoint.Set 0, 0, parSupportHeight + parVesselDiameter * 0.5
    Set ObjEndCap = PlaceRevolution(m_OutputColl, ellipse1, revVect, centerPoint, PI * 2, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEndCap
    
' Set the output
    
    Set ObjEndCap = Nothing
    Set Objcylinder = Nothing
    Set revVect = Nothing
    Set centerPoint = Nothing
    Set geomFactory = Nothing
    
    'Delete the closed curve which was created for the revolution.
    Dim gscadElem As IJDObject
    Set gscadElem = Linecurve
    Debug.Assert Not gscadElem Is Nothing
    Set Linecurve = Nothing
    gscadElem.Remove
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
