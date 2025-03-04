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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Tuesday, Nov 16 1999
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India) Copyright Information, Header  is added.
'  08.SEP.2006     KKC                DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.APR.2007     KKC                CR-117893 Modified Olet symbol to be Trimmable
'   2.Jul.2007     svsmylav           TR-123363: Added PI/2 value to parAngle to support V6.1 instances.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI           As Double
Private RAD          As Double
Private gscadElem    As IJDObject

Private Const MODULE = "Physical:" 'Used for error messages


Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Dim parBranchFittingPlacementType As Long
        
    Set oPartFclt = arrayOfInputs(1)
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
     
    parBranchFittingPlacementType = arrayOfInputs(7)
    
    If lPartDataBasis <= 1 Then
        If parBranchFittingPlacementType = 5 Then
            Call runPDB1(m_OutputColl, arrayOfInputs, arrayOfOutputs)
        Else
            Call runPDB1onFitting(m_OutputColl, arrayOfInputs, arrayOfOutputs)
        End If
    End If
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
Private Sub runPDB1(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

Const METHOD = "runPDB1"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim Objbody As Object
    
    Dim parA As Double
    Dim parB As Double
    Dim parC As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parB = arrayOfInputs(3)
    parC = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    If UBound(arrayOfInputs) >= 6 Then 'Not to fail the prior to V7 instances by accessing index 6
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 2 'To suppport V6.1 instances
    End If
    
    iOutput = 0

' Insert your code for output 1(olet body)
    Dim HeaderOD As Double 'Header Outside pipe diameter
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, HeaderOD, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    'get branch data
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim stPoint         As New AutoMath.DPosition
    
    
    'do some calcs to get the point positions
    Dim p1 As New AutoMath.DPosition
    Dim p2 As New AutoMath.DPosition
    Dim p3 As New AutoMath.DPosition
    Dim p4 As New AutoMath.DPosition
    Dim O2Rad As Double
    
    Dim L As Double
    Dim H1 As Double
    Dim H2 As Double
    Dim dTmp As Double
    
    'get the larger of the flange diam or pipe diam for the output diam
    If CmpDblLessThan(flangeDiam, pipeDiam) Then
        O2Rad = pipeDiam / 2
    Else
        O2Rad = flangeDiam / 2
        If CmpDblLessThan(parB, flangeDiam) Then
        parB = parB * 1.2
        Else
        parB = parB
        End If
    End If
    
    'set the start point in the center of the pipe
    Dim verysmalldistance As Double
    verysmalldistance = 0.0001
    p1.Set verysmalldistance, 0, 0
    
    L = (parB - parC) / 2#
    H1 = L / Tan(40# / RAD)
    
'    p2.Set parB / 2, H1, 0
    p2.Set parB / 2, H1 + HeaderOD / 2, 0
    
'    p3.Set O2Rad, parA, 0
    p3.Set O2Rad, parA + HeaderOD / 2, 0
    
    Set lines = New Collection
    
    Set oLine = PlaceTrLine(p1, p2)
    lines.Add oLine
    Set oLine = PlaceTrLine(p2, p3)
    lines.Add oLine
    
    Set objCStr = PlaceTrCString(p1, lines)
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
    
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
    
'    centPoint.Set 0, 0, 0
    centPoint.Set 0, 0 + HeaderOD / 2, 0
    axis.Set 0, 1, 0
    Set Objbody = PlaceRevolution(m_OutputColl, objCStr, _
                                    axis, centPoint, 2 * PI, True)
    Set objCStr = Nothing
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    Objbody.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), Objbody
    Set Objbody = Nothing
    Set oDirVector = Nothing
    Set oTransformationMat = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double
    Dim HeaderRadius As Double
    
    HeaderRadius = pipeDiam / 2

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)

    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                       flangeThick, flangeDiam, cptOffset, depth
                       
    oPlacePoint.Set (parA + HeaderRadius + cptOffset - depth) * Cos(parAngle), (parA + HeaderRadius + cptOffset - depth) * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
Private Sub runPDB1onFitting(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

Const METHOD = "runPDB1"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim Objbody As Object
    
    Dim parA As Double
    Dim parB As Double
    Dim parC As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parB = arrayOfInputs(3)
    parC = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    If UBound(arrayOfInputs) >= 6 Then 'Not to fail the prior to V7 instances by accessing index 6
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 2 'To suppport V6.1 instances
    End If
    
    iOutput = 0

' Insert your code for output 1(olet body)
    Dim HeaderOD As Double 'Header Outside pipe diameter
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, HeaderOD, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    'get branch data
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim stPoint         As New AutoMath.DPosition
    
    
    'do some calcs to get the point positions
    Dim p1 As New AutoMath.DPosition
    Dim p2 As New AutoMath.DPosition
    Dim p3 As New AutoMath.DPosition
    Dim p4 As New AutoMath.DPosition
    Dim O2Rad As Double
    
    Dim L As Double
    Dim H1 As Double
    Dim H2 As Double
    Dim dTmp As Double
    
    'get the larger of the flange diam or pipe diam for the output diam
    If CmpDblLessThan(flangeDiam, pipeDiam) Then
        O2Rad = pipeDiam / 2
    Else
        O2Rad = flangeDiam / 2
        If CmpDblLessThan(parB, flangeDiam) Then
        parB = parB * 1.2
        Else
        parB = parB
        End If
    End If
    
    'set the start point in the center of the pipe
    Dim verysmalldistance As Double
    verysmalldistance = 0.0001
    p1.Set verysmalldistance, -parA / 2, 0
    p2.Set O2Rad, -parA / 2, 0
        
    L = (parB - parC) / 2#
    H1 = L / Tan(40# / RAD)

    p3.Set parB / 2, -parA / 2 + H1, 0
    
'    p3.Set O2Rad, parA, 0
    p4.Set O2Rad, parA / 2, 0
    
    Set lines = New Collection
    
    Set oLine = PlaceTrLine(p1, p2)
    lines.Add oLine
    Set oLine = PlaceTrLine(p2, p3)
    lines.Add oLine
    Set oLine = PlaceTrLine(p3, p4)
    lines.Add oLine
    
    Set objCStr = PlaceTrCString(p1, lines)
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
    
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
    
'    centPoint.Set 0, 0, 0
    centPoint.Set 0, -parA / 2, 0
    axis.Set 0, 1, 0
    Set Objbody = PlaceRevolution(m_OutputColl, objCStr, _
                                    axis, centPoint, 2 * PI, True)
    Set objCStr = Nothing
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    Objbody.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), Objbody
    Set Objbody = Nothing
    Set oDirVector = Nothing
    Set oTransformationMat = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double
    Dim HeaderRadius As Double
    
    HeaderRadius = pipeDiam / 2

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)

    oPlacePoint.Set -(parA / 2) * Cos(parAngle), -(parA / 2) * Sin(parAngle), 0
    oDir.Set -Cos(parAngle), -Sin(parAngle), 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Place Nozzle 2
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                       flangeThick, flangeDiam, cptOffset, depth
                       
    oPlacePoint.Set (parA / 2 + cptOffset - depth) * Cos(parAngle), (parA / 2 + cptOffset - depth) * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0

    Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    Dim oPipePort1 As IJDPipePort
    Set oPipePort1 = objNozzle1

    Dim oPipePort2 As IJDPipePort
    Set oPipePort2 = objNozzle2
    
    oPipePort1.NPD = oPipePort2.NPD
    oPipePort1.NPDUnitType = oPipePort2.NPDUnitType

' Set the output (Nozzle 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing
    
    
' Set the output (Nozzle 2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    
    Set oPipePort1 = Nothing
    Set oPipePort2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


