VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 1 2004
'   Description:
'     This Symbol details were taken from PDS Piping Component Data SN=OP_572 Page # D-142
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
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

    Dim iOutput     As Double
    Dim ObjMaintInclinedStem As Object
    Dim ObjMaintRectOperator As Object
    
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
   

' Insert your code for output 3(Maintenance Stem of Operator)
    startPoint.Set 0, 0, 0
    endPoint.x = -Cos(parStemAngle) * (parOperatorHeight - parOperatorDepth)
    endPoint.y = Sin(parStemAngle) * (parOperatorHeight - parOperatorDepth)
    endPoint.z = 0
    Set ObjMaintInclinedStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintInclinedStem
    Set ObjMaintInclinedStem = Nothing
    
' Insert your code for output 4(Maintenance Rectangle Prism )
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
    Set ObjMaintRectOperator = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        oLineString, _
                                                        ProjectionPoint.x, ProjectionPoint.y, ProjectionPoint.z, _
                                                        -parOperatorDepth, isCapped)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintRectOperator
    Set ObjMaintRectOperator = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set ProjectionPoint = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing
     
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
