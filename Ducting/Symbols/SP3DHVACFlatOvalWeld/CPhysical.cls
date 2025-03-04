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
'   Author:          MS
'   Creation Date:  December 10 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' The symbol consists of 6 outputs: Four physical outputs and  Two hvac ports
' The weld cross-section is taken to be semi-elliptical.
' In which 2 physical outputs are created by 'PlaceRevolution' and the other 2 by 'PlaceProjection'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.DEC.2004     MS     CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'   01.Aug.2006     KKK       Replaced 'm_OutputColl.ResourceManager' with "Nothing'.

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
    Dim ObjTopWeldSeam As Object
    Dim ObjBottomWeldSeam As Object
    Dim ObjFrontWeldSeam As Object
    Dim ObjRearWeldSeam As Object
    
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
'   CreateHvacNozzle' method requires non-zero width and depth values.
'   Check if parDepth is ZERO and set its value to parWidth.
   If parDepth = 0 Then parDepth = parWidth
    

    iOutput = 0
    
 'Remove this incorrect adjustment of the Width; Route should fix the duct definition to match the port command spec
' FlatOval:Modified value of parWidth will be used by 'CreateHvacNozzle'
''''    parWidth = parWidth - parDepth
    Dim projDist       As Double
    projDist = parWidth - parDepth
    
    If projDist < NEGLIGIBLE_THICKNESS Then projDist = NEGLIGIBLE_THICKNESS
    
' Weld thicknes is taken as 20mm, and weld height is taken as 5 mm
'   these can be adjusted as required
    WeldThickness = 0.02
    WeldHeight = 0.005

' Prepare ellipse in X-Y plane which forms weld seam cross-section
    Dim CenterX   As Double
    Dim CenterY   As Double
    Dim CenterZ   As Double
    Dim ellipseMajorX   As Double
    Dim ellipseMajorY   As Double
    Dim ellipseMajorZ   As Double
    Dim ellipseMMRatio  As Double
    Dim ObjEllipticalArc As IngrGeom3D.EllipticalArc3d
   
    ellipseMMRatio = WeldHeight / (WeldThickness / 2)
    CenterX = 0
    CenterY = parDepth / 2
    CenterZ = 0
    ellipseMajorX = WeldThickness / 2
    ellipseMajorY = 0
    ellipseMajorZ = 0
    Set ObjEllipticalArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, CenterX, CenterY, CenterZ, _
                                            0, 0, 1, ellipseMajorX, ellipseMajorY, ellipseMajorZ, _
                                            ellipseMMRatio, 0, PI)

'Used for transformation of weld cross-section.
    Dim transMat        As New AutoMath.DT4x4
    Dim transVec        As New AutoMath.DVector
        
' Insert your code for output 1(Front Weld Seam)
    ' Transform the ellipse arc so as to move by projDist/2
    transMat.LoadIdentity
    transVec.Set 0, 0, projDist / 2
    transMat.Translate transVec
    ObjEllipticalArc.Transform transMat
        
    ' Revolution at Point 1
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, 0, projDist / 2
    Set ObjFrontWeldSeam = PlaceRevolution(m_OutputColl, ObjEllipticalArc, axisVect, centerPoint, PI, True)
        
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontWeldSeam
    Set ObjFrontWeldSeam = Nothing
        
' Insert your code for output 2(Top Weld Seam)
    ' Projection at Point 1
    axisVect.Set 0, 0, -1
    Set ObjTopWeldSeam = PlaceProjection(m_OutputColl, ObjEllipticalArc, axisVect, projDist, True)
         
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopWeldSeam
    Set ObjTopWeldSeam = Nothing
    
' Insert your code for output 3(Rear Weld Seam)
    ' Revolution at point2
    transMat.LoadIdentity
    transVec.Set 0, 0, -projDist
    transMat.Translate transVec
    ObjEllipticalArc.Transform transMat
    axisVect.Set -1, 0, 0
    centerPoint.Set 0, 0, -projDist / 2
    Set ObjRearWeldSeam = PlaceRevolution(m_OutputColl, ObjEllipticalArc, axisVect, centerPoint, PI, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearWeldSeam
    Set ObjRearWeldSeam = Nothing
    
' Insert your code for output 4(Bottom Weld Seam)
    '  Projection at Point 4
    transMat.LoadIdentity
    axisVect.Set -1, 0, 0
    transMat.Rotate PI, axisVect
    ObjEllipticalArc.Transform transMat
    axisVect.Set 0, 0, -1
    Set ObjBottomWeldSeam = PlaceProjection(m_OutputColl, ObjEllipticalArc, axisVect, projDist, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomWeldSeam
    Set ObjBottomWeldSeam = Nothing

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

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.FlatOval, EndPrep, _
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
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(2, "SymbDefn", GSCADNozzleEntities.FlatOval, EndPrep, _
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
    Set transVec = Nothing
    
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = ObjEllipticalArc
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set ObjEllipticalArc = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
    
