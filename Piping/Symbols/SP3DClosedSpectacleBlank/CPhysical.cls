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
'   Creation Date:  Thursday, Jan 25 2001
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
    Dim ObjInlineBlank As Object
    Dim ObjInterCyl As Object
    Dim ObjOpenDisc As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parFig8InterAxis As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFig8InterAxis = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 1(In line Blank Disc)
    Dim parDiscDiameter As Double
    parDiscDiameter = parFig8InterAxis * 0.8
    Dim Port1Position   As New AutoMath.DPosition
    Port1Position.Set -parFacetoFace / 2, 0, 0
    Dim Port2Position   As New AutoMath.DPosition
    Port2Position.Set parFacetoFace / 2, 0, 0

    Set ObjInlineBlank = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, parDiscDiameter, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlineBlank
    Set ObjInlineBlank = Nothing
    
 ' Insert your code for output 2(Cylinder bet Discs)
    Dim parInterDiscsDiameter As Double
    parInterDiscsDiameter = parFacetoFace / 2
    Dim CylStPosition   As New AutoMath.DPosition
    CylStPosition.Set 0, parFig8InterAxis * 0.3, 0
    Dim CylEndPosition   As New AutoMath.DPosition
    CylEndPosition.Set 0, parFig8InterAxis * 0.7, 0
    Set ObjInterCyl = PlaceCylinder(m_OutputColl, CylStPosition, CylEndPosition, parInterDiscsDiameter, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInterCyl
    Set ObjInterCyl = Nothing
    
 ' Insert your code for output 3(Open Disc)
    Dim axis        As New AutoMath.DVector
    axis.Set 1, 0, 0
    Dim OpenDiscAxisCentPosition   As New AutoMath.DPosition
    OpenDiscAxisCentPosition.Set 0, parFig8InterAxis, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parPipetoDisc As Double
    parPipetoDisc = (parDiscDiameter - pipeDiam) / 2
    Dim parHubThick As Double
    parHubThick = parPipetoDisc * 0.2
    
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d

    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, -parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, -parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0, parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0)
    lines.Add oLine
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0, parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0)
    lines.Add oLine
    
    
    Dim stPoint   As New AutoMath.DPosition
    stPoint.Set -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0
    Set objCStr = PlaceTrCString(stPoint, lines)
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
'''    ......     use our complex string (e.g. for projection or revolution)
    Set ObjOpenDisc = PlaceRevolution(m_OutputColl, objCStr, axis, OpenDiscAxisCentPosition, 3.141586 * 2, False)
    Set objCStr = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOpenDisc
    Set ObjOpenDisc = Nothing
    
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
