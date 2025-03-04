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
'   Author:          NN
'   Creation Date:  Tuesday, Feb 6 2001
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjHandleFace As Plane3d
    Dim ObjHandleBack As Plane3d
    Dim ObjHandleSide As Object
    Dim ObjHandleHole As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parHandletoCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parHandletoCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 1(Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parDiscDiameter As Double
    parDiscDiameter = pipeDiam * 1.24
    Dim Port1Position   As New AutoMath.DPosition
    Port1Position.Set -parFacetoFace / 2, 0, 0
    Dim Port2Position   As New AutoMath.DPosition
    Port2Position.Set parFacetoFace / 2, 0, 0

    Set ObjBody = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, parDiscDiameter, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
 ' Insert your code for output 2(Face of Handle)
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dblPnts(0 To 11)       As Double
    dblPnts(0) = parFacetoFace / 4
    dblPnts(1) = parDiscDiameter * 0.4
    dblPnts(2) = -parFacetoFace * 0.75

    dblPnts(3) = parFacetoFace / 4
    dblPnts(4) = parDiscDiameter * 0.4
    dblPnts(5) = parFacetoFace * 0.75

    dblPnts(6) = parFacetoFace / 4
    dblPnts(7) = parHandletoCenter
    dblPnts(8) = parFacetoFace * 0.75

    dblPnts(9) = parFacetoFace / 4
    dblPnts(10) = parHandletoCenter
    dblPnts(11) = -parFacetoFace * 0.75
    
    Dim nn0        As Double
    Dim nn1        As Double
    Dim nn2        As Double
    Dim nn3        As Double
    Dim nn4        As Double
    Dim nn5        As Double
    Dim nn6        As Double
    Dim nn7        As Double
    Dim nn8        As Double
    Dim nn9        As Double
    Dim nn10       As Double
    Dim nn11       As Double

    nn0 = parFacetoFace / 4
    nn1 = parHandletoCenter - 1.5 * parFacetoFace
    nn2 = parFacetoFace / 2

    nn3 = parFacetoFace / 4
    nn4 = parHandletoCenter - 1.5 * parFacetoFace
    nn5 = -parFacetoFace / 2
    
    nn6 = parFacetoFace / 4
    nn7 = parHandletoCenter - 0.5 * parFacetoFace
    nn8 = -parFacetoFace / 2

    nn9 = parFacetoFace / 4
    nn10 = parHandletoCenter - 0.5 * parFacetoFace
    nn11 = parFacetoFace / 2
            
    Dim parNumberPoints As Long
    parNumberPoints = 4
    Set ObjHandleFace = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, parNumberPoints, dblPnts)
    
    Set geomFactory = Nothing
    
    Dim objHoleF    As IngrGeom3D.ComplexString3d
    
    Dim onnGeomFactory    As IngrGeom3D.GeometryFactory
    Set onnGeomFactory = New IngrGeom3D.GeometryFactory
    Dim nnstartPoint   As New AutoMath.DPosition
    Dim objArc1       As IngrGeom3D.Arc3d
    Dim nnArcs           As Collection
    Set nnArcs = New Collection

    Set objArc1 = onnGeomFactory.Arcs3d.CreateBy3Points(Nothing, nn0, nn1, 0, nn0, parHandletoCenter - parFacetoFace, nn5, nn0, nn7, 0)
    nnArcs.Add objArc1
    
    Set objArc1 = onnGeomFactory.Arcs3d.CreateBy3Points(Nothing, nn0, nn7, 0, nn0, parHandletoCenter - parFacetoFace, nn2, nn0, nn1, 0)
    nnArcs.Add objArc1
    
    nnstartPoint.Set nn0, nn1, 0

    Set objHoleF = PlaceTrCString(nnstartPoint, nnArcs)
    
    ObjHandleFace.AddBoundary objHoleF

    Set onnGeomFactory = Nothing
    Dim iCount As Integer
    For iCount = 1 To nnArcs.Count
        nnArcs.Remove 1
    Next iCount
    Set nnArcs = Nothing
    Set objHoleF = Nothing
    Set objArc1 = Nothing

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleFace
    Set ObjHandleFace = Nothing
    
 ' Insert your code for output 3(Back of Handle)

    Dim BdblPnts(0 To 11)       As Double
    BdblPnts(0) = -parFacetoFace / 4
    BdblPnts(1) = parDiscDiameter * 0.4
    BdblPnts(2) = -parFacetoFace * 0.75
    
    BdblPnts(3) = -parFacetoFace / 4
    BdblPnts(4) = parDiscDiameter * 0.4
    BdblPnts(5) = parFacetoFace * 0.75
    
    BdblPnts(6) = -parFacetoFace / 4
    BdblPnts(7) = parHandletoCenter
    BdblPnts(8) = parFacetoFace * 0.75
    
    BdblPnts(9) = -parFacetoFace / 4
    BdblPnts(10) = parHandletoCenter
    BdblPnts(11) = -parFacetoFace * 0.75
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set ObjHandleBack = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, parNumberPoints, BdblPnts)
    Set geomFactory = Nothing
    Dim oBnnGeomFactory     As IngrGeom3D.GeometryFactory
    Dim objHoleB            As IngrGeom3D.ComplexString3d

    Set oBnnGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim objArc2       As IngrGeom3D.Arc3d
    Dim nnArcs2           As Collection
    Set nnArcs2 = New Collection

    Set objArc2 = oBnnGeomFactory.Arcs3d.CreateBy3Points(Nothing, -nn0, nn1, 0, -nn0, parHandletoCenter - parFacetoFace, nn5, -nn0, nn7, 0)
    nnArcs2.Add objArc2
    
    Set objArc2 = oBnnGeomFactory.Arcs3d.CreateBy3Points(Nothing, -nn0, nn7, 0, -nn0, parHandletoCenter - parFacetoFace, nn2, -nn0, nn1, 0)
    nnArcs2.Add objArc2
        
    Dim BnnstartPoint   As New AutoMath.DPosition
    BnnstartPoint.Set -nn0, nn1, 0
    Set objHoleB = PlaceTrCString(BnnstartPoint, nnArcs2)
        
    ObjHandleBack.AddBoundary objHoleB

    Set oBnnGeomFactory = Nothing
    Dim iCount2 As Integer
    For iCount2 = 1 To nnArcs2.Count
        nnArcs2.Remove 1
    Next iCount2
    Set nnArcs2 = Nothing
    Set objHoleB = Nothing
    Set objArc2 = Nothing
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleBack
    Set ObjHandleBack = Nothing
    
 ' Insert your code for output 4(Sides of Handle)
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d

    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BdblPnts(0), BdblPnts(1), BdblPnts(2), BdblPnts(3), BdblPnts(4), BdblPnts(5))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BdblPnts(3), BdblPnts(4), BdblPnts(5), BdblPnts(6), BdblPnts(7), BdblPnts(8))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BdblPnts(6), BdblPnts(7), BdblPnts(8), BdblPnts(9), BdblPnts(10), BdblPnts(11))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BdblPnts(9), BdblPnts(10), BdblPnts(11), BdblPnts(0), BdblPnts(1), BdblPnts(2))
    lines.Add oLine
    
    Dim startPoint   As New AutoMath.DPosition
    startPoint.Set BdblPnts(0), BdblPnts(1), BdblPnts(2)
    Set objCStr = PlaceTrCString(startPoint, lines)
    Set oLine = Nothing
    Dim iCountSide As Integer
    For iCountSide = 1 To lines.Count
        lines.Remove 1
    Next iCountSide
    Set lines = Nothing
    Dim axis        As New AutoMath.DVector
    axis.Set 1, 0, 0
    Set ObjHandleSide = PlaceProjection(m_OutputColl, objCStr, axis, parFacetoFace * 0.5, False)
    Set objCStr = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleSide
    Set ObjHandleSide = Nothing
    
 ' Insert your code for output 5(Hole in Handle)
    Dim HoleDiam        As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -parFacetoFace / 4, parHandletoCenter - parFacetoFace, 0
    enPoint.Set parFacetoFace / 4, parHandletoCenter - parFacetoFace, 0
    HoleDiam = parFacetoFace
    Set ObjHandleHole = PlaceCylinder(m_OutputColl, stPoint, enPoint, HoleDiam, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleHole
    Set ObjHandleHole = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim FacetoFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    FacetoFace = arrayOfInputs(2)
    oPlacePoint.Set -FacetoFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set FacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim cptoffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - cptoffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + cptoffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
