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
'   Copyright(c)2005-06 Intergraph GmbH
'   All Rights Reserved
'
'  "SimplePhysical:".cls
'  Author:          svsmylav/KKK
'   Creation Date:  Dienstag, Jul 5 2005
'  Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      TODO - fill in header description information
'
'  Change History:
'  dd.mmm.yyyy    who    change description
'  -----------    ---    ------------------
'
'  08.SEP.2006     KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.SEP.2006     RRK/DKL TR-104419: Removed input parameter 'Depth'
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
 
    Dim iOutput     As Double
        
'    {<(InputParamDec)>}
    Dim Width As Double
    Dim Depth As Double
'    Dim BWidth As Double
'    Dim BDepth As Double
'    Dim Length1 As Double
'    Dim Length2 As Double
'    Dim E1 As Double
'    Dim E2 As Double
'    Dim Direction As Double
    
    Dim blnOpposite As Boolean
 
    Dim TroxBEZ As String
    Dim DiffuserSide As Long
    Dim DiffuserHeight As Double
    Dim DiffuserWidth As Double
    Dim dblToInner As Double
    Dim dblToDown As Double
       
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    DiffuserHeight = arrayOfInputs(3)
    DiffuserWidth = arrayOfInputs(4)
    DiffuserSide = arrayOfInputs(5)
    
    Depth = Width
    dblToInner = 0.005
    dblToDown = 0.01
    
    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0
    
' Insert your code for output 1(Quader)

    Dim ObjBodyA As Object
    Dim pPos1 As New AutoMath.DPosition
    Dim pPos2 As New AutoMath.DPosition
    
    Dim objP1 As AutoMath.DPosition
    Dim objP2 As AutoMath.DPosition
    Dim objP3 As AutoMath.DPosition
    Dim objP4 As AutoMath.DPosition
    
    If DiffuserSide = 0 Or DiffuserSide = 2 Then
    
        pPos1.Set -DiffuserWidth / 2#, -DiffuserWidth / 2#, -DiffuserHeight / 2#
        pPos2.Set DiffuserWidth / 2#, DiffuserWidth / 2#, DiffuserHeight / 2#
        
    Else
    
        pPos1.Set -DiffuserWidth / 2#, -DiffuserHeight / 2#, -DiffuserWidth / 2#
        pPos2.Set DiffuserWidth / 2#, DiffuserHeight / 2#, DiffuserWidth / 2#
        
    End If
    Set ObjBodyA = PlaceBox(m_OutputColl, pPos1, pPos2)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyA
    Set ObjBodyA = Nothing
    
' Place circle 1
    Dim dblDirection As Double
    
    If DiffuserSide = 0 Or DiffuserSide = 1 Then
        dblDirection = 1#
    Else
        dblDirection = -1#
    End If
    
    Dim objTruncCylinder  As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim Height1     As Double
    Dim Height2     As Double
    Dim Dir As New AutoMath.DVector
    Dim Dir1 As New AutoMath.DVector
    Dim Dir2 As New AutoMath.DVector
    Dim DirDown As New AutoMath.DVector
    Dim Rotation As Double
'    Dim Diam As Double
    Dim Size1 As Double
    Dim Size2 As Double
    
    
    If DiffuserSide = 0 Or DiffuserSide = 2 Then
        stPoint.Set 0, 0, dblDirection * (-DiffuserHeight / 2# + 0.003)
        Dir.Set 0, 0, -1
        
        Dir1.Set 1, 0, 0
        Dir2.Set 0, 1, 0
        Size1 = DiffuserWidth
        Size2 = DiffuserWidth
        DirDown.Set 0, 0, -dblDirection * dblToDown
    Else
        stPoint.Set 0, dblDirection * (-DiffuserHeight / 2# + 0.003), 0
        Dir.Set 0, -1, 0
        
        Dir1.Set 1, 0, 0
        Dir2.Set 0, 0, 1
        Size1 = DiffuserWidth
        Size2 = DiffuserWidth
        DirDown.Set 0, -dblDirection * dblToDown, 0
    End If
    
    Height1 = dblDirection * 0.006
    Height2 = dblDirection * 0.006
    
    Rotation = 0
'    Diam = TroxDP
'    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(m_OutputColl, stPoint, Height1, Height2, Dir, Rotation, Diam, False)
'
''    Set objTruncCylinder = PlaceFAN(m_OutputColl, stPoint, 0.16)
'
'    iOutput = iOutput + 1
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTruncCylinder
'    Set objTruncCylinder = Nothing
    
    
 
    
    
' Place circle 2
 
    If DiffuserSide = 0 Or DiffuserSide = 2 Then
        stPoint.Set 0, 0, dblDirection * (-DiffuserHeight / 2# + 0.003)
        Dir.Set 0, 0, -1
    Else
        stPoint.Set 0, dblDirection * (-DiffuserHeight / 2# + 0.003), 0
        Dir.Set 0, -1, 0
    End If
    Height1 = dblDirection * 0.006
    Height2 = dblDirection * 0.006
    
    Rotation = 0
'    Diam = TroxDP / 3#
'    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(m_OutputColl, stPoint, Height1, Height2, Dir, Rotation, Diam, False)
'    iOutput = iOutput + 1
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTruncCylinder
'    Set objTruncCylinder = Nothing
    
    
    
    
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
    Set Dir = New AutoMath.DVector
    Dim RadialDir As New AutoMath.DVector
    Dim CornerRadius As Double
        
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

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
    pos.Set -DiffuserWidth / 2# - 0.035, 0, 0
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector Dir
    
    Dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient Dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 0.035
                                                             
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle


    ' add the code for the lüfter
    Dim lngIndex As Long
    Dim i As Long
    Dim dblDist As Double
    
    lngIndex = 0
    dblDist = 0.02
    
    For i = 0 To lngNRippen - 1
    
    Set objP1 = vecDir2(stPoint, Dir1, 0.5 * Size1 - i * dblDist, Dir2, 0.5 * Size2 - i * dblDist)
    Set objP2 = vecDir2(stPoint, Dir1, -0.5 * Size1 + i * dblDist, Dir2, 0.5 * Size2 - i * dblDist)
    Set objP3 = vecDir2(stPoint, Dir1, -0.5 * Size1 + i * dblDist, Dir2, -0.5 * Size2 + i * dblDist)
    Set objP4 = vecDir2(stPoint, Dir1, 0.5 * Size1 - i * dblDist, Dir2, -0.5 * Size2 + i * dblDist)
    Call createRect(m_OutputColl, "Luft", lngIndex, objP1, objP2, objP3, objP4, dblToInner, DirDown)


    Next i
    
    Set iDistribPort = Nothing

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub
    Resume
    
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
'  DirDown is the size and direction in down position
'


' Compute the 4 inner Points
Dim objVec1 As AutoMath.DVector
 
 

Dim objP1I As AutoMath.DPosition
Dim objP2I As AutoMath.DPosition
Dim objP3I As AutoMath.DPosition
Dim objP4I As AutoMath.DPosition

 

Set objVec1 = objP3.Subtract(objP1)
objVec1.Length = dblToInner
Set objP1I = objP1.Offset(objVec1)

Set objVec1 = objP4.Subtract(objP2)
objVec1.Length = dblToInner
Set objP2I = objP2.Offset(objVec1)

Set objVec1 = objP1.Subtract(objP3)
objVec1.Length = dblToInner
Set objP3I = objP3.Offset(objVec1)

Set objVec1 = objP2.Subtract(objP4)
objVec1.Length = dblToInner
Set objP4I = objP4.Offset(objVec1)

 

' Compute the 4 curves (which will later be projected)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP1, objP2, objP2I, objP1I, DirDown)
        
'        Exit Function
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
                    objVec As AutoMath.DVector) As Long
                    
    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP1.x, objP1.y, objP1.z, objP2.x, objP2.y, objP2.z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP2.x, objP2.y, objP2.z, objP3.x, objP3.y, objP3.z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.x, objP3.y, objP3.z, objP4.x, objP4.y, objP4.z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.x, objP4.y, objP4.z, objP1.x, objP1.y, objP1.z)
    complex.AddCurve oLine, True
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.x, objVec.y, objVec.z, objVec.Length, True)
                                                    
    lngIndex = lngIndex + 1
    objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
     
     
End Function
                    
                    
Private Function vecDir2(Pin As AutoMath.DPosition, Dir1 As AutoMath.DVector, dblSize1 As Double, _
                                                    Dir2 As AutoMath.DVector, dblSize2 As Double) As AutoMath.DPosition

Dim Dir1X As AutoMath.DVector
Dim Dir2X As AutoMath.DVector

Set Dir1X = Dir1.Clone
Dir1X.Length = dblSize1

Set Dir2X = Dir2.Clone
Dir2X.Length = dblSize2

Set vecDir2 = Pin.Offset(Dir1X)
Set vecDir2 = vecDir2.Offset(Dir2X)


End Function

