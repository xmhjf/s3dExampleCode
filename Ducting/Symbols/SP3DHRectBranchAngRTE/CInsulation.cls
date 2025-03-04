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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Monday, Jan 13 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation: 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjHvacHeaderIns As Object
    Dim ObjHvacBranchIns As Object
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHLength As Double
    Dim parBLength As Double
'   Angle is measured from positive X-axis in anti-clockwise direction
    Dim parAngle As Double
    Dim parInsulationThickness As Double

    Dim ProjVector As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 14) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHLength = arrayOfInputs(6)
    parBLength = arrayOfInputs(7)
    parAngle = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0
    
    Dim minHL As Double
    Dim minBL As Double
    Dim halfHLength As Double
    Dim HD As Double
    Dim HW As Double
    
'   Define default minimum Tee or Reducing Tee header and branch lengths
    minHL = parWidth + 0.05
    minBL = minHL / 2
    If parHLength < minHL Then
        parHLength = minHL
    End If
    If parBLength < minBL Then
        parBLength = minBL
    End If
    halfHLength = parHLength / 2
    
    HD = parDepth / 2
    HW = parWidth / 2

' Insert your code for output 4(Hvac Header Insulation)

'   Point 1
    LineStringPoints(0) = -halfHLength
    LineStringPoints(1) = HD + parInsulationThickness
    LineStringPoints(2) = HW + parInsulationThickness
    
'   Point 2
    LineStringPoints(3) = -halfHLength
    LineStringPoints(4) = HD + parInsulationThickness
    LineStringPoints(5) = -(HW + parInsulationThickness)
    
'   Point 3
    LineStringPoints(6) = -halfHLength
    LineStringPoints(7) = -(HD + parInsulationThickness)
    LineStringPoints(8) = -(HW + parInsulationThickness)
    
'   Point 4
    LineStringPoints(9) = -halfHLength
    LineStringPoints(10) = -(HD + parInsulationThickness)
    LineStringPoints(11) = HW + parInsulationThickness
    
'   Point 5
    LineStringPoints(12) = LineStringPoints(0)
    LineStringPoints(13) = LineStringPoints(1)
    LineStringPoints(14) = LineStringPoints(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStringPoints)
    ProjVector.Set 1, 0, 0
    Set ObjHvacHeaderIns = PlaceProjection(m_OutputColl, oLineString, ProjVector, parHLength, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacHeaderIns
    Set ObjHvacHeaderIns = Nothing
    
' Insert your code for output 5(Hvac Branch Insulation)
'   Create line string at the Origin (0, 0, 0)

'   Update half depth and half width values for the branch
    HD = parBDepth / 2
    HW = parBWidth / 2
    
'   Point 1
    LineStringPoints(0) = -(HD + parInsulationThickness) * Sin(parAngle)
    LineStringPoints(1) = (HD + parInsulationThickness) * Cos(parAngle)
    LineStringPoints(2) = HW + parInsulationThickness
    
'   Point 2
    LineStringPoints(3) = -(HD + parInsulationThickness) * Sin(parAngle)
    LineStringPoints(4) = (HD + parInsulationThickness) * Cos(parAngle)
    LineStringPoints(5) = -(HW + parInsulationThickness)
    
'   Point 3
    LineStringPoints(6) = (HD + parInsulationThickness) * Sin(parAngle)
    LineStringPoints(7) = -(HD + parInsulationThickness) * Cos(parAngle)
    LineStringPoints(8) = -(HW + parInsulationThickness)
    
'   Point 4
    LineStringPoints(9) = (HD + parInsulationThickness) * Sin(parAngle)
    LineStringPoints(10) = -(HD + parInsulationThickness) * Cos(parAngle)
    LineStringPoints(11) = HW + parInsulationThickness
    
'   Point 5
    LineStringPoints(12) = LineStringPoints(0)
    LineStringPoints(13) = LineStringPoints(1)
    LineStringPoints(14) = LineStringPoints(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStringPoints)
    ProjVector.Set Cos(parAngle), Sin(parAngle), 0
    Set ObjHvacBranchIns = PlaceProjection(m_OutputColl, oLineString, ProjVector, parBLength, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacBranchIns
    Set ObjHvacBranchIns = Nothing
    
    Set ProjVector = Nothing
 
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
 
    Set geomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
