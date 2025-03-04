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
'   Copyright(c)2005 Intergraph GmbH
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
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   03.Nov.2005     kkk     CR-87366  Create TROX HVAC parts and symbols
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
    
 
    Dim iOutput     As Double
        
'    {<(InputParamDec)>}
    Dim Width As Double
    Dim Depth As Double
 
    Dim blnOpposite As Boolean
    Dim FinWidth As Double
    Dim FrameWidth As Double
    Dim FrameThickness As Double
    Dim GrillWidth As Double
    Dim GrillLength As Double
    
    Dim dblToInner As Double
    Dim dblToDown As Double
    Dim dblNozzleLength As Double
       
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    FinWidth = arrayOfInputs(4)
    FrameWidth = arrayOfInputs(5)
    FrameThickness = arrayOfInputs(6)
    GrillWidth = arrayOfInputs(7)
    GrillLength = arrayOfInputs(8)
    
    dblToInner = 0.005
    dblToDown = 0.01
    
    dblNozzleLength = 0.001 ' 0.035
    
    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0
    
' Insert your code for output 1(Nozzle)



    
    Dim objP1 As AutoMath.DPosition
    Dim objP2 As AutoMath.DPosition
    Dim objP3 As AutoMath.DPosition
    Dim objP4 As AutoMath.DPosition
    Dim stPoint   As New AutoMath.DPosition
    Dim Dir1 As New AutoMath.DVector
    Dim Dir2 As New AutoMath.DVector
    Dim Dir3 As New AutoMath.DVector
    Dim lngIndex As Long
    Dim DirDown As New AutoMath.DVector
 
    Dim dblA As Double
    Dim dblHeight As Double
    
    lngIndex = 0
    
    stPoint.Set 0, 0, 0
    
    Dir1.Set 1, 0, 0
    Dir2.Set 0, 1, 0
    Dir3.Set 0, 0, 1

    DirDown.Set 0, 0, FrameThickness
    
    Set objP1 = vecDir2(stPoint, Dir1, 0.5 * GrillLength + FrameWidth, Dir2, 0.5 * GrillWidth + FrameWidth)
    Set objP2 = vecDir2(stPoint, Dir1, -0.5 * GrillLength - FrameWidth, Dir2, 0.5 * GrillWidth + FrameWidth)
    Set objP3 = vecDir2(stPoint, Dir1, -0.5 * GrillLength - FrameWidth, Dir2, -0.5 * GrillWidth - FrameWidth)
    Set objP4 = vecDir2(stPoint, Dir1, 0.5 * GrillLength + FrameWidth, Dir2, -0.5 * GrillWidth - FrameWidth)
    Call createRect(m_OutputColl, "Feature", lngIndex, objP1, objP2, objP3, objP4, FrameWidth, DirDown)
    
    

    DirDown.Set -GrillLength, 0, 0
     
    
    dblA = Sqr(0.5 * FrameThickness * FrameThickness)
    dblHeight = 0
    
    While dblHeight < GrillWidth - FinWidth
    
        Set objP1 = vecDir2(stPoint, Dir1, 0.5 * GrillLength, Dir2, 0.5 * GrillWidth - dblHeight, Dir3, dblA)
        Set objP2 = vecDir2(stPoint, Dir1, 0.5 * GrillLength, Dir2, 0.5 * GrillWidth - dblA - dblHeight, Dir3, 0)
        Set objP3 = vecDir2(stPoint, Dir1, 0.5 * GrillLength, Dir2, 0.5 * GrillWidth - FinWidth - dblHeight, Dir3, FrameThickness - dblA)
        Set objP4 = vecDir2(stPoint, Dir1, 0.5 * GrillLength, Dir2, 0.5 * GrillWidth - FinWidth + dblA - dblHeight, Dir3, FrameThickness)
        Call createProjectedLines(m_OutputColl, "Feature", lngIndex, _
                                        objP1, objP2, objP3, objP4, DirDown)
            
        dblHeight = dblHeight + FinWidth * 1.2
    
    Wend
    
    Debug.Print "Maximum index: " & lngIndex
 

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
                                                    complex, objVec.x, objVec.y, objVec.z, objVec.Length, dblClosed)
                                                    
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



