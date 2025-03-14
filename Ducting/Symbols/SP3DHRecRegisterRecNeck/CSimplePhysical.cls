VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'   Definition of HVAC Rectangular Register. The symbol graphics and data are extracted from
'   Carnes Catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private Const E_FAIL = &H80004005
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
      PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CP As New AutoMath.DPosition
    Dim PortDirection As New AutoMath.DVector
    Dim RadialDirection As New AutoMath.DVector
    Dim Width As Double
    Dim Depth As Double
    Dim iOutput     As Double
    Dim GrillBladeWidth As Double
    Dim DamperBladeWidth As Double
    Dim FrameWidth As Double
    Dim RegisterThickness As Double
    Dim RegisterWidth As Double
    Dim RegisterLength As Double
    Dim dblToInner As Double
    Dim dblToDown As Double
    Dim NeckLength As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2) 'Duct Width
    Depth = arrayOfInputs(3) 'Duct Depth
    GrillBladeWidth = arrayOfInputs(4) ' Blade Width of Inner Grills
    DamperBladeWidth = arrayOfInputs(5) ' Blade Width of Outer Damper Blades
    FrameWidth = arrayOfInputs(6) ' Stiffness Frame Width - 1 Inch
    RegisterThickness = arrayOfInputs(7)  ' Register Thickness
    RegisterWidth = arrayOfInputs(8) 'Register Width ' The Height of Register including Frame Width
    RegisterLength = arrayOfInputs(9) ' Register Length = Total Length of the Register Including Frame Width
    NeckLength = arrayOfInputs(10) ' Neck Length
    
    dblToInner = 0.005
    dblToDown = 0.01
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    iOutput = 0
'=====================================
'BUILD HVAC NOZZLE: Nozzle - OutPut - 1
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim cptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    Dim FlangeWidth As Double
    Dim CPos       As New AutoMath.DPosition
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim NozzleLength As Double, Thickness As Double
'Set HVAC nozzle parameters
    iPortIndex = 1
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
    PortDepth = 0#
    cptOffset = 0#
    
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlowDir = DistribFlow_IN

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HvacNozzle1", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos
    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir

    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength
' Set the output
    m_OutputColl.AddOutput "HvacNozzle1", oHvacNozzle
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CPos = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    
' Insert your code for Output 2 (Rectangular Neck)
Dim point1   As New AutoMath.DPosition
Dim point2 As New AutoMath.DPosition
Dim ObjRegisterNeck As Object

point1.Set 0, -Depth / 2, -Width / 2
point2.Set NeckLength / 2, Depth / 2, Width / 2

Set ObjRegisterNeck = PlaceBox(m_OutputColl, point1, point2)
m_OutputColl.AddOutput "Neck", ObjRegisterNeck

' Insert your code for Output 2 (Conical Portion)

Dim ObjRegisterCone As IngrGeom3D.RuledSurface3d
point1.Set NeckLength / 2, 0, 0

Set ObjRegisterCone = PlaceTrapezoid(m_OutputColl, point1, Width, Depth, RegisterLength, RegisterWidth, NeckLength / 2, False, 0, PI / 2, 0)
m_OutputColl.AddOutput "ConePortion", ObjRegisterCone

' Insert your code for output 1(Grille Body)

    Dim objP1 As AutoMath.DPosition
    Dim objP2 As AutoMath.DPosition
    Dim objP3 As AutoMath.DPosition
    Dim objP4 As AutoMath.DPosition
    Dim stPoint   As New AutoMath.DPosition
    Dim stPoint2   As New AutoMath.DPosition
    Dim Dir1 As New AutoMath.DVector
    Dim Dir2 As New AutoMath.DVector
    Dim Dir3 As New AutoMath.DVector
    Dim lngIndex As Long
    Dim DirDown As New AutoMath.DVector
 
    Dim dblA As Double
    Dim dblHeight As Double
     
    lngIndex = 4
    stPoint.Set NeckLength, 0, 0
    stPoint2.Set NeckLength + RegisterThickness / 2, 0, 0
    
    Dir1.Set 1, 0, 0
    Dir2.Set 0, 1, 0
    Dir3.Set 0, 0, 1
    
    DirDown.Set RegisterThickness / 2, 0, 0
    
    Set objP1 = vecDir2(stPoint, Dir3, 0.5 * RegisterLength, Dir2, 0.5 * RegisterWidth)
    Set objP2 = vecDir2(stPoint, Dir3, -0.5 * RegisterLength, Dir2, 0.5 * RegisterWidth)
    Set objP3 = vecDir2(stPoint, Dir3, -0.5 * RegisterLength, Dir2, -0.5 * RegisterWidth)
    Set objP4 = vecDir2(stPoint, Dir3, 0.5 * RegisterLength, Dir2, -0.5 * RegisterWidth)
    Call createRect(m_OutputColl, "Feature", lngIndex, objP1, objP2, objP3, objP4, FrameWidth, DirDown)
        
' Now create Deflection Blades of Grille
    
    DirDown.Set 0, 0, -(RegisterLength - 2 * FrameWidth)
    
    dblA = Sqr(0.5 * RegisterThickness / 2 * RegisterThickness / 2)
    dblHeight = 0
    
    While CmpDblLessThan(dblHeight, ((RegisterWidth - 2 * FrameWidth) - GrillBladeWidth))
    
        Set objP1 = vecDir2(stPoint, Dir3, 0.5 * (RegisterLength - 2 * FrameWidth), Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth) - dblHeight, Dir1, dblA)
        Set objP2 = vecDir2(stPoint, Dir3, 0.5 * (RegisterLength - 2 * FrameWidth), Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth) - dblA - dblHeight, Dir1, 0)
        Set objP3 = vecDir2(stPoint, Dir3, 0.5 * (RegisterLength - 2 * FrameWidth), Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth) - GrillBladeWidth - dblHeight, Dir1, RegisterThickness / 2 - dblA)
        Set objP4 = vecDir2(stPoint, Dir3, 0.5 * (RegisterLength - 2 * FrameWidth), Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth) - GrillBladeWidth + dblA - dblHeight, Dir1, RegisterThickness / 2)
        Call createProjectedLines(m_OutputColl, "Feature", lngIndex, _
                                        objP1, objP2, objP3, objP4, DirDown)
            
        dblHeight = dblHeight + GrillBladeWidth
    
    Wend
'Insert Your Code for Damper above the Grille

    DirDown.Set RegisterThickness / 2, 0, 0
    
    Set objP1 = vecDir2(stPoint2, Dir3, 0.5 * RegisterLength, Dir2, 0.5 * RegisterWidth)
    Set objP2 = vecDir2(stPoint2, Dir3, -0.5 * RegisterLength, Dir2, 0.5 * RegisterWidth)
    Set objP3 = vecDir2(stPoint2, Dir3, -0.5 * RegisterLength, Dir2, -0.5 * RegisterWidth)
    Set objP4 = vecDir2(stPoint2, Dir3, 0.5 * RegisterLength, Dir2, -0.5 * RegisterWidth)
    Call createRect(m_OutputColl, "Feature", lngIndex, objP1, objP2, objP3, objP4, FrameWidth, DirDown)
    
' Now create Damper Baffle Blades
    
    DirDown.Set 0, -(RegisterWidth - 2 * FrameWidth), 0
    
    dblA = Sqr(0.5 * RegisterThickness / 2 * RegisterThickness / 2)
    dblHeight = 0
    
    While CmpDblLessThan(dblHeight, (RegisterLength - 2 * FrameWidth - DamperBladeWidth))
    
        Set objP1 = vecDir2(stPoint2, Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth), Dir3, 0.5 * (RegisterLength - 2 * FrameWidth) - dblHeight, Dir1, dblA)
        Set objP2 = vecDir2(stPoint2, Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth), Dir3, 0.5 * (RegisterLength - 2 * FrameWidth) - dblA - dblHeight, Dir1, 0)
        Set objP3 = vecDir2(stPoint2, Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth), Dir3, 0.5 * (RegisterLength - 2 * FrameWidth) - DamperBladeWidth - dblHeight, Dir1, RegisterThickness / 2 - dblA)
        Set objP4 = vecDir2(stPoint2, Dir2, 0.5 * (RegisterWidth - 2 * FrameWidth), Dir3, 0.5 * (RegisterLength - 2 * FrameWidth) - DamperBladeWidth + dblA - dblHeight, Dir1, RegisterThickness / 2)
        Call createProjectedLines(m_OutputColl, "Feature", lngIndex, _
                                        objP1, objP2, objP3, objP4, DirDown)
            
        dblHeight = dblHeight + DamperBladeWidth
    
    Wend
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
    Set m_GeomFactory = Nothing
End Sub
Private Function createRect(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    dblToInner As Double, _
                    DirDown As AutoMath.DVector) As Long
                    
' create 4 quader resulting in a rectangle:
'               1             2
'               +-------------+
'               |             |
'               |  +-------+  |
'               |  |       |  |
'               |  |       |  |
'               |  |       |  |
'               |  +-------+  |
'               |             |
'               +-------------+
'               4             3
'
'  P1 - P4 are the 4 outer corner points
'  dblToInner is the size from the outer points to the inner points
'  The size is measured for each side.
'  DirDown is the size and direction in down position
'


' Compute the 4 inner Points
Dim objVec1 As AutoMath.DVector
Dim objVec2 As AutoMath.DVector
 

Dim objP1I As AutoMath.DPosition
Dim objP2I As AutoMath.DPosition
Dim objP3I As AutoMath.DPosition
Dim objP4I As AutoMath.DPosition

 

Set objVec1 = objP2.Subtract(objP1)
Set objVec2 = objP4.Subtract(objP1)
Set objP1I = vecDir2(objP1, objVec1, dblToInner, objVec2, dblToInner)
 
Set objVec1 = objP1.Subtract(objP2)
Set objVec2 = objP3.Subtract(objP2)
Set objP2I = vecDir2(objP2, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP2.Subtract(objP3)
Set objVec2 = objP4.Subtract(objP3)
Set objP3I = vecDir2(objP3, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP3.Subtract(objP4)
Set objVec2 = objP1.Subtract(objP4)
Set objP4I = vecDir2(objP4, objVec1, dblToInner, objVec2, dblToInner)

' Compute the 4 curves (which will later be projected)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP1, objP2, objP2I, objP1I, DirDown)
        
         
        
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP2, objP3, objP3I, objP2I, DirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP3, objP4, objP4I, objP3I, DirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP4, objP1, objP1I, objP4I, DirDown)

End Function
Private Function createProjectedLines(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    objVec As AutoMath.DVector, _
                    Optional dblClosed As Boolean = True) As Long
                    
    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP1.X, objP1.Y, objP1.Z, objP2.X, objP2.Y, objP2.Z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP2.X, objP2.Y, objP2.Z, objP3.X, objP3.Y, objP3.Z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.X, objP3.Y, objP3.Z, objP4.X, objP4.Y, objP4.Z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.X, objP4.Y, objP4.Z, objP1.X, objP1.Y, objP1.Z)
    complex.AddCurve oLine, True
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.X, objVec.Y, objVec.Z, objVec.Length, dblClosed)
                                                    
    lngIndex = lngIndex + 1
    objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
     
     
End Function
                    
                    
Private Function vecDir2(Pin As AutoMath.DPosition, _
                         Dir1 As AutoMath.DVector, dblSize1 As Double, _
                         Optional Dir2 As AutoMath.DVector = Nothing, _
                            Optional dblSize2 As Double = 0#, _
                         Optional Dir3 As AutoMath.DVector = Nothing, _
                            Optional dblSize3 As Double = 0#) As AutoMath.DPosition

Dim Dir1X As AutoMath.DVector
Dim Dir2X As AutoMath.DVector

Set Dir1X = Dir1.Clone
Dir1X.Length = dblSize1
Set vecDir2 = Pin.Offset(Dir1X)

If Not Dir2 Is Nothing Then
    Set Dir2X = Dir2.Clone
    Dir2X.Length = dblSize2
    Set vecDir2 = vecDir2.Offset(Dir2X)
End If
If Not Dir3 Is Nothing Then
    Set Dir2X = Dir3.Clone
    Dir2X.Length = dblSize3
    Set vecDir2 = vecDir2.Offset(Dir2X)
End If
End Function
