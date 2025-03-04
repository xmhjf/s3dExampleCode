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
'   Author:          ACM
'   Creation Date:  Friday, Nov 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Gorunding Paddle symbol which is taken from www.resistoflex.com web catalog.
'   This symbol has 8 outputs, 7 physical and one Insulation.
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
    Dim ObjPaddleBody As Object
    Dim ObjHandleFrontPlane As Plane3d
    Dim ObjHandleBackPlane As Plane3d
    Dim ObjProjForHandleSides As Object
    Dim ObjHandleThroughHole As Object
    
    Dim parFacetoFace As Double
    Dim parPaddleHeight As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parPaddleHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 1(PaddleBody)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim PaddleDia As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    ''Paddle dia is taken as Flange dia at CP1
    PaddleDia = flangeDiam

    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0

    Set ObjPaddleBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, PaddleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPaddleBody
    Set ObjPaddleBody = Nothing
      
' Insert your code for output 4(Sides of Handle)

    Dim objCollection           As Collection
    Set objCollection = New Collection
    
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim HdlGeomCstr         As IngrGeom3D.ComplexString3d
    
    Dim oLine  As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oDir        As New AutoMath.DVector
    Dim PlanePoints(0 To 11) As Double
    Dim MeetPoint As Double
    Dim HoleDia As Double
    
    HoleDia = 0.009375  ''3/8" dia hole
    If HoleDia > parPaddleHeight / 4 Then HoleDia = parPaddleHeight / 4
     
    MeetPoint = Sqr((PaddleDia / 2) ^ 2 - (parPaddleHeight / 4) ^ 2)
    
    ''Points to create Handle Geometry
    PlanePoints(0) = parFacetoFace / 2
    PlanePoints(1) = MeetPoint
    PlanePoints(2) = parPaddleHeight / 4

    PlanePoints(3) = parFacetoFace / 2
    PlanePoints(4) = MeetPoint
    PlanePoints(5) = -parPaddleHeight / 4

    PlanePoints(6) = parFacetoFace / 2
    PlanePoints(7) = (PaddleDia / 2) + 0.75 * parPaddleHeight
    PlanePoints(8) = -parPaddleHeight / 4

    PlanePoints(9) = parFacetoFace / 2
    PlanePoints(10) = (PaddleDia / 2) + parPaddleHeight
    PlanePoints(11) = 0
   
   ''This is to create the sides of Handle, for which a complex string of front face is
   ''created and projected without caps
   
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, PlanePoints(0), PlanePoints(1), PlanePoints(2), _
                                                    PlanePoints(3), PlanePoints(4), PlanePoints(5))
    objCollection.Add oLine
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, PlanePoints(3), PlanePoints(4), PlanePoints(5), _
                                                    PlanePoints(6), PlanePoints(7), PlanePoints(8))
    objCollection.Add oLine
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, PlanePoints(6), PlanePoints(7), _
                                    PlanePoints(8), PlanePoints(9), PlanePoints(10), PlanePoints(11), _
                                    PlanePoints(6), PlanePoints(7), -PlanePoints(8))
    objCollection.Add oArc
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, PlanePoints(6), PlanePoints(7), -PlanePoints(8), _
                                                    PlanePoints(0), PlanePoints(1), PlanePoints(2))
    objCollection.Add oLine
   
    stPoint.Set PlanePoints(0), PlanePoints(1), PlanePoints(2)
    
    Set HdlGeomCstr = PlaceTrCString(stPoint, objCollection)
    
    oDir.Set -1, 0, 0
    Set ObjProjForHandleSides = PlaceProjection(m_OutputColl, HdlGeomCstr, oDir, parFacetoFace, False)
   
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProjForHandleSides
    Set ObjProjForHandleSides = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To objCollection.Count
        objCollection.Remove 1
    Next iCount
    
 ' Insert your code for output 2(Face of Handle)
    
    ''Handle Front Plane is prepared using the Handle Geometry Complex String
    Set ObjHandleFrontPlane = oGeomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, _
                                                                               HdlGeomCstr)
  
    
    Dim ThroughHole    As IngrGeom3D.ComplexString3d
    Dim ArcPoints(0 To 11) As Double
     
    ''These are Through Hole Points which is created by two arcs.
    ArcPoints(0) = parFacetoFace / 2
    ArcPoints(1) = PaddleDia / 2 + (0.75 * parPaddleHeight - HoleDia / 2)
    ArcPoints(2) = 0

    ArcPoints(3) = parFacetoFace / 2
    ArcPoints(4) = PaddleDia / 2 + (0.75 * parPaddleHeight)
    ArcPoints(5) = -HoleDia / 2
    
    ArcPoints(6) = parFacetoFace / 2
    ArcPoints(7) = PaddleDia / 2 + (0.75 * parPaddleHeight + HoleDia / 2)
    ArcPoints(8) = 0

    ArcPoints(9) = parFacetoFace / 2
    ArcPoints(10) = PaddleDia / 2 + (0.75 * parPaddleHeight)
    ArcPoints(11) = HoleDia / 2
     

'' This is regarding through hole.
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, ArcPoints(0), ArcPoints(1), _
                                    ArcPoints(2), ArcPoints(3), ArcPoints(4), ArcPoints(5), _
                                    ArcPoints(6), ArcPoints(7), ArcPoints(8))
    
    objCollection.Add oArc
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, ArcPoints(6), ArcPoints(7), _
                                    ArcPoints(8), ArcPoints(9), ArcPoints(10), ArcPoints(11), _
                                    ArcPoints(0), ArcPoints(1), ArcPoints(2))

    objCollection.Add oArc
    
    stPoint.Set ArcPoints(0), ArcPoints(1), ArcPoints(2)

    Set ThroughHole = PlaceTrCString(stPoint, objCollection)
    
    ''The Hole is added to plane as a boundary.
    ObjHandleFrontPlane.AddBoundary ThroughHole

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleFrontPlane
    Set ObjHandleFrontPlane = Nothing
       
    For iCount = 1 To objCollection.Count
        objCollection.Remove 1
    Next iCount
   
 ' Insert your code for output 3(Back of Handle)
      
    ''Creation of Outer Boundary for Back plance of the paddle Handle.
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -PlanePoints(0), PlanePoints(1), PlanePoints(2), _
                                                    -PlanePoints(3), PlanePoints(4), PlanePoints(5))
    objCollection.Add oLine
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -PlanePoints(3), PlanePoints(4), PlanePoints(5), _
                                                    -PlanePoints(6), PlanePoints(7), PlanePoints(8))
    objCollection.Add oLine
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -PlanePoints(6), PlanePoints(7), _
                                    PlanePoints(8), -PlanePoints(9), PlanePoints(10), PlanePoints(11), _
                                    -PlanePoints(6), PlanePoints(7), -PlanePoints(8))
    objCollection.Add oArc
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -PlanePoints(6), PlanePoints(7), -PlanePoints(8), _
                                                    -PlanePoints(0), PlanePoints(1), PlanePoints(2))
    objCollection.Add oLine
    
    stPoint.Set -PlanePoints(0), PlanePoints(1), PlanePoints(2)
    
    Set HdlGeomCstr = PlaceTrCString(stPoint, objCollection)
    
    Set ObjHandleBackPlane = oGeomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, _
                                                                                HdlGeomCstr)
       
     For iCount = 1 To objCollection.Count
        objCollection.Remove 1
     Next iCount
    
    '' Creation of through hole.
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -ArcPoints(0), ArcPoints(1), _
                                    ArcPoints(2), -ArcPoints(3), ArcPoints(4), ArcPoints(5), _
                                    -ArcPoints(6), ArcPoints(7), ArcPoints(8))
    
    objCollection.Add oArc
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -ArcPoints(6), ArcPoints(7), _
                                    ArcPoints(8), -ArcPoints(9), ArcPoints(10), ArcPoints(11), _
                                    -ArcPoints(0), ArcPoints(1), ArcPoints(2))
    
    objCollection.Add oArc
    
    stPoint.Set -ArcPoints(0), ArcPoints(1), ArcPoints(2)

    Set ThroughHole = PlaceTrCString(stPoint, objCollection)
    
    ''Adding Hole as boundary to back plane
    ObjHandleBackPlane.AddBoundary ThroughHole

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleBackPlane
    Set ObjHandleBackPlane = Nothing
        
    For iCount = 1 To objCollection.Count
        objCollection.Remove 1
    Next iCount
    
 ' Insert your code for output 5(Hole in Handle)
    
    ''It's a simple cylinder without caps.
    stPoint.Set parFacetoFace / 2, (PaddleDia / 2) + 0.75 * parPaddleHeight, 0
    enPoint.Set -parFacetoFace / 2, (PaddleDia / 2) + 0.75 * parPaddleHeight, 0
   
    Set ObjHandleThroughHole = PlaceCylinder(m_OutputColl, stPoint, enPoint, HoleDia, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleThroughHole
    Set ObjHandleThroughHole = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As New AutoMath.DPosition
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
   
    oPlacePoint.Set -parFacetoFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
   
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
   
   Set oPlacePoint = Nothing
   Set oDir = Nothing
   Set oLine = Nothing
   Set enPoint = Nothing
   Set stPoint = Nothing
   Set oArc = Nothing
   Set HdlGeomCstr = Nothing
   Set oGeomFactory = Nothing
   Set objCollection = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


