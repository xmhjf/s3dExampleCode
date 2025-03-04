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
'   Creation Date:  Sunday, March 3 2002
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
Private PI           As Double
Private RAD          As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

PI = 4 * Atn(1)
RAD = 180 / PI

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim ConduitDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim CenterPos       As New AutoMath.DPosition
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim iOutput     As Double
    Dim ObjHubport1 As Object
    Dim ObjBody As Object
    Dim ObjHubport2 As Object
    
    Dim parFacetoFace As Double
    Dim parUnionDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parUnionDiameter = arrayOfInputs(3)
    
    iOutput = 0
    CenterPos.Set 0, 0, 0

 ' Insert your code for output 1(Hub port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2
    
    Dim ConduitOD1 As Double
    ConduitOD1 = BodyOD(flangeDiam, ConduitDiam)
    Dim ConduitOD2 As Double
    ConduitOD2 = BodyOD(flangeDiam2, ConduitDiam2)
    Dim HubThickness1 As Double
    Dim HubThickness2 As Double
    HubThickness1 = 0.8 * flangeThick
    HubThickness2 = 0.8 * flangeThick2
    If flangeThick = 0 And flangeThick2 = 0 Then
    HubThickness1 = parFacetoFace / 3
    HubThickness2 = parFacetoFace / 3
        Else
            If flangeThick = 0 Then
            HubThickness1 = parFacetoFace - 2.2 * flangeThick2
            HubThickness2 = 0.8 * flangeThick2
            Else
                If flangeThick2 = 0 Then
                HubThickness2 = parFacetoFace - 2.2 * flangeThick
                HubThickness1 = 0.8 * flangeThick
            End If
        End If
    End If
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -(parFacetoFace - (HubThickness1 + HubThickness2)) / 2, 0, 0
    Set ObjHubport1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConduitOD1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHubport1
    Set ObjHubport1 = Nothing
 
 ' Insert your code for output 2(Body)
    Dim parA72rad           As Double
    Dim parHexaZ            As Double
    Dim parHexaY            As Double
    Dim parHexagoneA(21)     As Double
    parA72rad = PI / 3
    parHexaZ = parUnionDiameter / 2 * Cos(parA72rad)
    parHexaY = parUnionDiameter / 2 * Sin(parA72rad)
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
''    stPoint.Set -(parFacetoFace - (flangeThick + flangeThick2)) / 2, 0, 0
''    enPoint.Set (parFacetoFace - (flangeThick + flangeThick2)) / 2, 0, 0
    stPoint.Set -(parFacetoFace - (HubThickness1 + HubThickness2)) / 2, 0, 0
    enPoint.Set (parFacetoFace - (HubThickness1 + HubThickness2)) / 2, 0, 0

' Point 1
    parHexagoneA(0) = stPoint.x
    parHexagoneA(1) = 0
    parHexagoneA(2) = -parUnionDiameter / 2
' Point 2
    parHexagoneA(3) = stPoint.x
    parHexagoneA(4) = -parHexaY
    parHexagoneA(5) = -parHexaZ
' Point 3
    parHexagoneA(6) = stPoint.x
    parHexagoneA(7) = -parHexaY
    parHexagoneA(8) = parHexaZ
' Point 4
    parHexagoneA(9) = stPoint.x
    parHexagoneA(10) = 0
    parHexagoneA(11) = parUnionDiameter / 2
' Point 5
    parHexagoneA(12) = stPoint.x
    parHexagoneA(13) = parHexaY
    parHexagoneA(14) = parHexaZ
' Point 6
    parHexagoneA(15) = stPoint.x
    parHexagoneA(16) = parHexaY
    parHexagoneA(17) = -parHexaZ
' Point 7
    parHexagoneA(18) = stPoint.x
    parHexagoneA(19) = 0
    parHexagoneA(20) = -parUnionDiameter / 2
    
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d

    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parHexagoneA(0), parHexagoneA(1), parHexagoneA(2), parHexagoneA(3), parHexagoneA(4), parHexagoneA(5))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parHexagoneA(3), parHexagoneA(4), parHexagoneA(5), parHexagoneA(6), parHexagoneA(7), parHexagoneA(8))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parHexagoneA(6), parHexagoneA(7), parHexagoneA(8), parHexagoneA(9), parHexagoneA(10), parHexagoneA(11))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parHexagoneA(9), parHexagoneA(10), parHexagoneA(11), parHexagoneA(12), parHexagoneA(13), parHexagoneA(14))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parHexagoneA(12), parHexagoneA(13), parHexagoneA(14), parHexagoneA(15), parHexagoneA(16), parHexagoneA(17))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parHexagoneA(15), parHexagoneA(16), parHexagoneA(17), parHexagoneA(18), parHexagoneA(19), parHexagoneA(20))
    lines.Add oLine
    
    Dim parHexaLength As Double
    parHexaLength = parFacetoFace - (HubThickness1 + HubThickness2)
    Dim objHexagoneComplex As IngrGeom3D.ComplexString3d
    Dim nnstartPoint   As New AutoMath.DPosition
    nnstartPoint.Set parHexagoneA(0), parHexagoneA(1), parHexagoneA(2)
    
    Set objHexagoneComplex = PlaceTrCString(nnstartPoint, lines)
    
        Set oLine = Nothing
    Dim iCountSide As Integer
    For iCountSide = 1 To lines.Count
        lines.Remove 1
    Next iCountSide
    Set lines = Nothing
    Set oGeomFactory = Nothing
    
    Set ObjBody = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        objHexagoneComplex, 1, 0, 0, parHexaLength, True)
    Set objHexagoneComplex = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    Set geomFactory = Nothing
 
 ' Insert your code for output 3(Hub port 2)
    stPoint.Set (parFacetoFace - (HubThickness1 + HubThickness2)) / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    Set ObjHubport2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConduitOD2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHubport2
    Set ObjHubport2 = Nothing
     
' Place Nozzle 1
    
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJConduitPortOcc
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)
    oPlacePoint.Set -faceToFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set faceToFace / 2 + cptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 2)

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
