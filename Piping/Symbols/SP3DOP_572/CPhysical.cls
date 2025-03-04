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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Friday, Dec 21 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'					      that they donot appear for small valve sizes
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInclinedStem As Object
    Dim ObjRectOperator As Object
    
    Dim parStemAngle As Double
    Dim parOperatorHeight As Double
    Dim parOperatorLength As Double
    Dim parOperatorLength1 As Double
    Dim parOperatorWidth As Double
    Dim parOperatorDepth As Double
    Dim parStemDia As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStemAngle = arrayOfInputs(2)
    parOperatorHeight = arrayOfInputs(3)
    parOperatorLength = arrayOfInputs(4)
    parOperatorLength1 = arrayOfInputs(5)
    parOperatorWidth = arrayOfInputs(6)
    parOperatorDepth = arrayOfInputs(7)
    
    iOutput = 0
'    parStemDia = 0.05
    parStemDia = (parOperatorLength + parOperatorLength1) * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   

' Insert your code for output 1(Stem of Operator)
    startPoint.Set 0, 0, 0
    endPoint.x = -Cos(parStemAngle) * (parOperatorHeight - parOperatorDepth)
    endPoint.y = Sin(parStemAngle) * (parOperatorHeight - parOperatorDepth)
    endPoint.z = 0
    Set ObjInclinedStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInclinedStem
    Set ObjInclinedStem = Nothing
    
' Insert your code for output 2(Rectangle Prism )
'Construct Points
    Dim dblPnts(0 To 14)  As Double
    Dim ProjectionPoint As New AutoMath.DPosition
    Dim oLineString As IngrGeom3D.LineString3d
    Dim isCapped As Boolean
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    isCapped = True
'Lower Right Point
    dblPnts(0) = -((parOperatorHeight - parOperatorDepth) * Cos(parStemAngle)) - (parOperatorLength * Sin(parStemAngle))
    dblPnts(1) = (parOperatorHeight - parOperatorDepth) * Sin(parStemAngle) - (parOperatorLength * Cos(parStemAngle))
    dblPnts(2) = parOperatorWidth / 2
'Upper Right Point
    dblPnts(3) = (-Cos(parStemAngle) * (parOperatorHeight - parOperatorDepth)) + parOperatorLength1 * Sin(parStemAngle)
    dblPnts(4) = (Sin(parStemAngle) * (parOperatorHeight - parOperatorDepth)) + (parOperatorLength1 * Cos(parStemAngle))
    dblPnts(5) = parOperatorWidth / 2
'Upper Left Point
    dblPnts(6) = (-Cos(parStemAngle) * (parOperatorHeight - parOperatorDepth)) + parOperatorLength1 * Sin(parStemAngle)
    dblPnts(7) = (Sin(parStemAngle) * (parOperatorHeight - parOperatorDepth)) + (parOperatorLength1 * Cos(parStemAngle))
    dblPnts(8) = -parOperatorWidth / 2
'Lower Left Point
    dblPnts(9) = -((parOperatorHeight - parOperatorDepth) * Cos(parStemAngle)) - (parOperatorLength * Sin(parStemAngle))
    dblPnts(10) = (parOperatorHeight - parOperatorDepth) * Sin(parStemAngle) - (parOperatorLength * Cos(parStemAngle))
    dblPnts(11) = -parOperatorWidth / 2
'Lower Right Point
    dblPnts(12) = -((parOperatorHeight - parOperatorDepth) * Cos(parStemAngle)) - (parOperatorLength * Sin(parStemAngle))
    dblPnts(13) = (parOperatorHeight - parOperatorDepth) * Sin(parStemAngle) - (parOperatorLength * Cos(parStemAngle))
    dblPnts(14) = parOperatorWidth / 2
'Calculate Points for Projection
    startPoint.x = (dblPnts(12) + dblPnts(3)) / 2
    startPoint.y = (dblPnts(13) + dblPnts(7)) / 2
    startPoint.z = 0
    endPoint.x = (dblPnts(12) + dblPnts(3)) / 2 - parOperatorDepth * Cos(parStemAngle)
    endPoint.y = (dblPnts(13) + dblPnts(7)) / 2 + parOperatorDepth * Sin(parStemAngle)
    endPoint.z = 0
    ProjectionPoint.x = startPoint.x - endPoint.x
    ProjectionPoint.y = startPoint.y - endPoint.y
    ProjectionPoint.z = startPoint.z - endPoint.z

'Place Line String
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPnts)
'Project the Line String to the required Direction
    Set ObjRectOperator = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        oLineString, _
                                                        ProjectionPoint.x, ProjectionPoint.y, ProjectionPoint.z, _
                                                        -parOperatorDepth, isCapped)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRectOperator
    Set ObjRectOperator = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
