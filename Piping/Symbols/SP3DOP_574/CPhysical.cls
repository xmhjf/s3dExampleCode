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
'   Creation Date:  Sunday, Dec 23 2001
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
    Dim ObjStem As Object
    Dim ObjMainCylinder As Object
    Dim ObjEccentricShaft As Object
    Dim ObjRectBox As Object
    Dim ObjRightCylinder As Object
    Dim ObjLeftCylinder As Object
    
    Dim parStemAngle As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parActuatorOffset As Double
    Dim parArmLength As Double
    Dim parArmDiameter As Double
    Dim parArmLength1 As Double
    Dim parActuatorBoxLength As Double
    Dim parActuatorBoxLength1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorBoxWidth As Double
    Dim parActuatorBoxDepth1 As Double
    Dim parActuatorBoxDepth As Double
    Dim parActuatorCenter As Double
    Dim parActuatorDiameter As Double
    Dim parStemDia As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStemAngle = arrayOfInputs(2)
    parOperatorHeight = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    parOperCenFromTop = arrayOfInputs(5)
    parOperatorDepth = arrayOfInputs(6)
    parActuatorOffset = arrayOfInputs(7)
    parArmLength = arrayOfInputs(8)
    parArmDiameter = arrayOfInputs(9)
    parArmLength1 = arrayOfInputs(10)
    parActuatorBoxLength = arrayOfInputs(11)
    parActuatorBoxLength1 = arrayOfInputs(12)
    parActuatorLength = arrayOfInputs(13)
    parActuatorLength1 = arrayOfInputs(14)
    parActuatorBoxWidth = arrayOfInputs(15)
    parActuatorBoxDepth1 = arrayOfInputs(16)
    parActuatorBoxDepth = arrayOfInputs(17)
    parActuatorCenter = arrayOfInputs(18)
    parActuatorDiameter = arrayOfInputs(19)
    
    
    iOutput = 0
'    parStemDia = 0.05
     parStemDia = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   

 ' Insert your code for output 1(Vertical Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set -(parOperatorHeight + parOperCenFromTop - parOperatorDepth) * Cos(parStemAngle), _
                    (parOperatorHeight + parOperCenFromTop - parOperatorDepth) * Sin(parStemAngle), 0
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Horizontal Cylinder)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set -(parOperatorHeight + parOperCenFromTop) * Cos(parStemAngle), _
                        (parOperatorHeight + parOperCenFromTop) * Sin(parStemAngle), 0
    Set ObjMainCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainCylinder
    Set ObjMainCylinder = Nothing
    
 ' Insert your code for output 3(Cylinder)
    startPoint.Set -parOperatorHeight * Cos(parStemAngle), _
                        parOperatorHeight * Sin(parStemAngle), _
                        -parActuatorOffset
    endPoint.Set -parOperatorHeight * Cos(parStemAngle) + parArmLength * Sin(parStemAngle), _
                       parOperatorHeight * Sin(parStemAngle) + parArmLength * Cos(parStemAngle), _
                       -parActuatorOffset
    Set ObjEccentricShaft = PlaceCylinder(m_OutputColl, startPoint, endPoint, parArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEccentricShaft
    Set ObjEccentricShaft = Nothing
    
' Insert your code for output 4(Box)
    Dim dblPnts(0 To 14)  As Double
    Dim ProjectionPoint As New AutoMath.DPosition
    Dim oLineString As IngrGeom3D.LineString3d
    Dim isCapped As Boolean
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    isCapped = True
    'Lower Right Corner
    dblPnts(0) = -parOperatorHeight * Cos(parStemAngle) _
            + parArmLength * Sin(parStemAngle) + (parActuatorBoxDepth - parActuatorBoxDepth1) * Cos(parStemAngle)
    dblPnts(1) = parOperatorHeight * Sin(parStemAngle) _
            + parArmLength * Cos(parStemAngle) - (parActuatorBoxDepth - parActuatorBoxDepth1) * Sin(parStemAngle)
    dblPnts(2) = parActuatorBoxLength - parActuatorOffset
    'Lower Left Corner
    dblPnts(3) = dblPnts(0)
    dblPnts(4) = dblPnts(1)
    dblPnts(5) = -parActuatorOffset - parActuatorBoxLength1
    'UpperRight Corner
    dblPnts(9) = -parOperatorHeight * Cos(parStemAngle) _
                        + parArmLength * Sin(parStemAngle) _
                        + (parActuatorBoxDepth - parActuatorBoxDepth1) * Cos(parStemAngle) _
                        + parActuatorBoxWidth * Sin(parStemAngle)
    dblPnts(10) = parOperatorHeight * Sin(parStemAngle) _
                        + parArmLength * Cos(parStemAngle) _
                        - (parActuatorBoxDepth - parActuatorBoxDepth1) * Sin(parStemAngle) _
                        + parActuatorBoxWidth * Cos(parStemAngle)
    dblPnts(11) = parActuatorBoxLength - parActuatorOffset
    'UpperLeft Corner
    dblPnts(6) = dblPnts(9)
    dblPnts(7) = dblPnts(10)
    dblPnts(8) = -parActuatorOffset - parActuatorBoxLength1
    'Lower Right Corner
    dblPnts(12) = -parOperatorHeight * Cos(parStemAngle) _
                        + parArmLength * Sin(parStemAngle) _
                        + (parActuatorBoxDepth - parActuatorBoxDepth1) * Cos(parStemAngle)
    dblPnts(13) = parOperatorHeight * Sin(parStemAngle) _
                        + parArmLength * Cos(parStemAngle) _
                        - (parActuatorBoxDepth - parActuatorBoxDepth1) * Sin(parStemAngle)
    dblPnts(14) = parActuatorBoxLength - parActuatorOffset
    startPoint.x = (dblPnts(12) + dblPnts(9)) / 2
    startPoint.y = (dblPnts(13) + dblPnts(10)) / 2
    startPoint.z = (dblPnts(14) + dblPnts(5)) / 2
    endPoint.x = startPoint.x - parActuatorBoxDepth * Cos(parStemAngle)
    endPoint.y = startPoint.y + parActuatorBoxDepth * Sin(parStemAngle)
    endPoint.z = startPoint.z
    ProjectionPoint.x = endPoint.x - startPoint.x
    ProjectionPoint.y = endPoint.y - startPoint.y
    ProjectionPoint.z = endPoint.z - startPoint.z
'Place Line String Bottom of the Box
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblPnts)
'Project the Line String to the required Direction
    Set ObjRectBox = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        oLineString, _
                                                        ProjectionPoint.x, ProjectionPoint.y, ProjectionPoint.z, _
                                                        parActuatorBoxDepth, isCapped)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRectBox
    Set oLineString = Nothing
    Set ObjRectBox = Nothing
    
 ' Insert your code for output 5(Cylinder)
    Dim CylinderStart As New AutoMath.DPosition
    Dim CylinderEnd As New AutoMath.DPosition
    CylinderStart.x = -parOperatorHeight * Cos(parStemAngle) + parArmLength1 * Sin(parStemAngle) _
                            - parActuatorCenter * Cos(parStemAngle)
    CylinderStart.y = parOperatorHeight * Sin(parStemAngle) + parArmLength1 * Cos(parStemAngle) _
                            + parActuatorCenter * Sin(parStemAngle)
    CylinderStart.z = parActuatorLength - parActuatorOffset
    CylinderEnd.x = CylinderStart.x
    CylinderEnd.y = CylinderStart.y
    CylinderEnd.z = parActuatorBoxLength - parActuatorOffset
    Set ObjRightCylinder = PlaceCylinder(m_OutputColl, CylinderStart, CylinderEnd, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCylinder
    Set ObjRightCylinder = Nothing

 ' Insert your code for output 6(Left Cylinder)
    CylinderStart.z = -parActuatorOffset - parActuatorBoxLength1
    CylinderEnd.x = CylinderStart.x
    CylinderEnd.y = CylinderStart.y
    CylinderEnd.z = -parActuatorOffset - parActuatorLength1
    Set ObjLeftCylinder = PlaceCylinder(m_OutputColl, CylinderStart, CylinderEnd, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCylinder
    Set ObjLeftCylinder = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
