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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  December 10 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' HVAC Split Weld places weld seam for Rectangular ducts.
' Corners and co-ordinate system of the Rectangular shape are as follows:
'        Y
'        |
'    2-------1
'    |   |   |
'    |   O --|----> Z
'    |       |
'    3-------4
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjWeldSeam1 As Object
    Dim ObjWeldSeam2 As Object
    Dim ObjWeldSeam3 As Object
    Dim ObjWeldSeam4 As Object
    Dim ObjWeldSeam5 As Object
    Dim ObjWeldSeam6 As Object
    Dim ObjWeldSeam7 As Object
    Dim ObjWeldSeam8 As Object
    
    Dim parWidth As Double
    Dim parDepth As Double
    
    Dim WeldThickness As Double
    Dim WeldHeight As Double

    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim geomFactory As New IngrGeom3D.GeometryFactory
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    
    iOutput = 0
    
'   Weld thicknes is taken as 20mm, and weld height is taken as 5 mm
'   these can be adjusted as required
    WeldThickness = 0.02
    WeldHeight = 0.005

' Insert your code for output 1(Weld Seam1)
    Dim transMat        As New AutoMath.DT4x4   'Used for transformation of weld cross-section.
    Dim HW As Double    'Half width
    Dim HD As Double    'Half depth
    
    HW = parWidth / 2
    HD = parDepth / 2
    
'   Prepare ellipse in X-Y plane which forms weld seam cross-section
    Dim CenterX   As Double
    Dim CenterY   As Double
    Dim CenterZ   As Double
    Dim ellipseMajorX   As Double
    Dim ellipseMajorY   As Double
    Dim ellipseMajorZ   As Double
    Dim ellipseMMRatio  As Double
    Dim ObjEllipticalArc As IngrGeom3D.EllipticalArc3d
    Dim ObjEllipticalArc2 As IngrGeom3D.EllipticalArc3d
    Dim ObjEllipticalArc3 As IngrGeom3D.EllipticalArc3d
    Dim ObjEllipticalArc4 As IngrGeom3D.EllipticalArc3d
    
    ellipseMMRatio = WeldHeight / (WeldThickness / 2)
    CenterY = HD
    CenterX = 0
    CenterZ = HW
    ellipseMajorX = WeldThickness / 2
    ellipseMajorY = 0
    ellipseMajorZ = 0
    Set ObjEllipticalArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            m_OutputColl.ResourceManager, CenterX, CenterY, CenterZ, _
                                            0, 0, 1, ellipseMajorX, ellipseMajorY, ellipseMajorZ, _
                                            ellipseMMRatio, 0, PI)
    
'   Revolution at corner 1 (Arc start point at 1st corner is 1 and its end point is 1', same
'   notation is used for the other corners).
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, CenterY, CenterZ
    Set ObjWeldSeam1 = PlaceRevolution(m_OutputColl, ObjEllipticalArc, axisVect, centerPoint, PI / 2, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam1
    Set ObjWeldSeam1 = Nothing
    
' Insert your code for output 2(Weld Seam2)
'   Projection at Point 1'
    axisVect.Set 0, 0, -1
    Set ObjWeldSeam2 = PlaceProjection(m_OutputColl, ObjEllipticalArc, axisVect, parWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam2
    Set ObjWeldSeam2 = Nothing

' Insert your code for output 3(Weld Seam3)
'   Revolution at corner 2 - Prepare an Elliptical arc at 2'
    CenterY = HD
    CenterX = 0
    CenterZ = -HW
    Set ObjEllipticalArc2 = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                        m_OutputColl.ResourceManager, CenterX, CenterY, CenterZ, _
                                        0, 1, 0, ellipseMajorX, ellipseMajorY, ellipseMajorZ, _
                                        ellipseMMRatio, 0, PI)
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, HD, -HW
    Set ObjWeldSeam3 = PlaceRevolution(m_OutputColl, ObjEllipticalArc2, axisVect, centerPoint, PI / 2, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam3
    Set ObjWeldSeam3 = Nothing
    
' Insert your code for output 4(Weld Seam4)
'   Projection at Point 2
    axisVect.Set 0, -1, 0
    Set ObjWeldSeam4 = PlaceProjection(m_OutputColl, ObjEllipticalArc2, axisVect, parDepth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam4
    Set ObjWeldSeam4 = Nothing
    
' Insert your code for output 5(Weld Seam5)
'   Revolution at point3
    transMat.LoadIdentity
    axisVect.Set -1, 0, 0
    transMat.Rotate PI, axisVect    'After rotation point 1' becomes 3'
    ObjEllipticalArc.Transform transMat
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, -HD, -HW
    Set ObjWeldSeam5 = PlaceRevolution(m_OutputColl, ObjEllipticalArc, axisVect, centerPoint, PI / 2, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam5
    Set ObjWeldSeam5 = Nothing
            
' Insert your code for output 6(Weld Seam6)
'   Projection at Point 3'
    axisVect.Set 0, 0, 1 'From point 3' to point 4
    Set ObjWeldSeam6 = PlaceProjection(m_OutputColl, ObjEllipticalArc, axisVect, parWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam6
    Set ObjWeldSeam6 = Nothing
    
' Insert your code for output 7(Weld Seam7)
'   Revolution at point4
    transMat.LoadIdentity
    axisVect.Set -1, 0, 0
    transMat.Rotate PI, axisVect   'Transform Semi-Ellipsical section at point 2' to that at 4'
    ObjEllipticalArc2.Transform transMat
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, -HD, HW
    Set ObjWeldSeam7 = PlaceRevolution(m_OutputColl, ObjEllipticalArc2, axisVect, centerPoint, PI / 2, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam7
    Set ObjWeldSeam7 = Nothing
    
' Insert your code for output 8(Weld Seam8)
'   Projection at Point 4'
    axisVect.Set 0, 1, 0
    Set ObjWeldSeam8 = PlaceProjection(m_OutputColl, ObjEllipticalArc2, axisVect, parDepth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWeldSeam8
    Set ObjWeldSeam8 = Nothing

' Place Nozzle 1
    Dim FlangeWidth As Double
    Dim NozzleFlangeThickness As Double
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim RadialDir As New AutoMath.DVector
    Dim CornerRadius As Double
    Dim WeldGap As Double       'Distance between port1 and port2
        
'   Set HVAC nozzle parameters
    PortDepth = 0#

'   FlangeWidth and Flange thickness assumed to be negigible thickness, Nozzle length to be
'   greater than flange thickness in general.
    FlangeWidth = NEGLIGIBLE_THICKNESS
    NozzleFlangeThickness = NEGLIGIBLE_THICKNESS
    
    FlowDir = DistribFlow_UNDEFINED
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    EndPrep = 301
    CptOffset = 0
    CornerRadius = 0
    WeldGap = 0.004     'WeldGap is taken as 4mm and can be adjusted as required
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -WeldGap / 2
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

'   Direction specified here of the nozzle should be the direction in which pipe will be routed.
'   Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    RadialDir.Set 0, 1, 0
    iDistribPort.SetRadialOrient RadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 2 * NEGLIGIBLE_THICKNESS

    m_OutputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

' Place Nozzle 2
    EndPrep = 301
    FlowDir = DistribFlow_UNDEFINED
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(2, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac2", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
   
'   Position of the nozzle should be the conenct point of the nozzle
    Set iDistribPort = oHvacNozzle
    X = WeldGap / 2
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    iDistribPort.SetPortLocation pos

    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    iDistribPort.SetRadialOrient RadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 2 * NEGLIGIBLE_THICKNESS

    m_OutputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    
    Set iDistribPort = Nothing
    Set centerPoint = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    Set RadialDir = Nothing
    Set transMat = Nothing

    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = ObjEllipticalArc
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set ObjEllipticalArc = Nothing

    Set ObjRMCurve = ObjEllipticalArc2
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set ObjEllipticalArc2 = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
