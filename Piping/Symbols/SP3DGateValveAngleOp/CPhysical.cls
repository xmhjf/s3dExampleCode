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
'   Copyright (c) 2006-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'      Gate Valve with Angle Operator
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006      KKC                     DI-95670  Replace names with initials in all
'                                           revision history sheets and symbols
'  21.Feb.2007      RRK                     TR-113129 Used CmdblEqual functions for comparision of doubles
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


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), _
               arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim I               As Integer
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim Objo1       As Object
    Dim Objo2       As Object
    Dim Objo3       As Object
    
    Dim parValveWidth As Double
    Dim parValveHeight As Double
    Dim parHandwheelDiameter As Double
    Dim parHandwheelAngle As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveWidth = arrayOfInputs(2)
    parValveHeight = arrayOfInputs(3)
    parHandwheelDiameter = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    
    
    Dim VH          As Double
    Dim VW          As Double
    Dim HWD         As Double
    Dim HWA         As Double
    Dim NPD         As Double
    Dim FTK         As Double
    Dim fd          As Double
    Dim lines       As Collection
    Dim oLine       As IngrGeom3D.Line3d
    Dim oArc        As IngrGeom3D.Arc3d
    Dim objCStr     As IngrGeom3D.ComplexString3d
    Dim stPoint     As New AutoMath.DPosition
    Dim enPoint     As New AutoMath.DPosition
    Dim ldiam       As Double
    Dim objCylinder As Object
    
    VH = parValveHeight
    VW = parValveWidth
    HWD = parHandwheelDiameter
    
    'HWA = parHandwheelAngle / RAD
    'we are getting the angle in radians only
    HWA = parHandwheelAngle
    
    iOutput = 0
    
 ' Get the first nozzle data
 RetrieveParameters 1, oPartFclt, m_OutputColl, _
                    pipeDiam, flangeThick, flangeDiam, _
                    sptOffset, depth
    NPD = pipeDiam
    
    If CmpDblEqual(flangeThick, 0) Then
        FTK = NPD / 6
    Else
        FTK = flangeThick
    End If
    
    If CmpDblEqual(flangeDiam, 0) Then
        fd = NPD
    Else
        fd = flangeDiam
    End If
    

 ' Insert your code for output 1() valve body
 Dim pnts(5) As New AutoMath.DPosition
 
 Dim objBspline  As IngrGeom3D.BSplineCurve3d
 Dim dtmp As Double
 dtmp = parValveWidth / 2 - flangeThick
 pnts(0).Set -dtmp, 0, pipeDiam / 2
 pnts(1).Set -dtmp * 0.7, 0, pnts(0).z
 pnts(2).Set -dtmp / 2, 0, fd * 0.6
 pnts(3).Set dtmp / 2, 0, fd * 0.6
 pnts(4).Set dtmp * 0.7, 0, pnts(0).z
 pnts(5).Set dtmp, 0, pipeDiam / 2
 Set objBspline = PlaceTrBspline(3, pnts)

 Dim centPoint   As New AutoMath.DPosition
 Dim axis        As New AutoMath.DVector
 Dim objRevolution  As Object
 centPoint.Set 0, 0, 0
 axis.Set 1, 0, 0
 Set objRevolution = PlaceRevolution(m_OutputColl, objBspline, _
                                    axis, centPoint, 2 * PI, _
                                    False)

 'Delete the curve that was created for the revolution
 Set gscadElem = objBspline
 Set objBspline = Nothing
 gscadElem.Remove

 ' Set the output
 iOutput = iOutput + 1
 m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
 Set objRevolution = Nothing


' Insert your code for output 2() create the verticle bspline body
    Dim pnts1(1 To 10) As New AutoMath.DPosition
    
    pnts1(1).Set 0, FTK, NPD / 2
    pnts1(2).Set 0, fd / 2 + FTK, NPD / 2
    pnts1(3).Set 0, fd / 2 + FTK, fd / 2
    pnts1(4).Set 0, fd / 2 + FTK * 3, fd / 2
    pnts1(5).Set 0, fd / 2 + FTK * 3, NPD / 2
    pnts1(6).Set 0, fd / 2 + FTK * 3 + NPD / 2, FTK / 2
    pnts1(7).Set 0, fd / 2 + FTK * 3, 0
    
    Set lines = New Collection
    For I = 1 To 4
        Set oLine = PlaceTrLine(pnts1(I), pnts1(I + 1))
        lines.Add oLine
    Next I
   
    Dim nvec As New AutoMath.DVector
    nvec.Set -1, 0, 0
    Set oArc = PlaceTrArcByCenterNorm(pnts1(5), pnts1(6), pnts1(7), nvec)
    lines.Add oArc
    
    Set objCStr = PlaceTrCStringNoCheck(lines)
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing

    centPoint.Set 0, 0, 0
    axis.Set 0, 1, 0
    Set objRevolution = PlaceRevolution(m_OutputColl, objCStr, _
                                    axis, centPoint, 2 * PI, _
                                    False)
    
    Set objCStr = Nothing

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objRevolution = Nothing
    
    
    
    
' Insert your code for output 3() create the vertical screw
    stPoint.Set 0, 0, 0
    enPoint.Set 0, VH - HWD, 0
    ldiam = parHandwheelDiameter / 15
    
    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, _
                                    enPoint, ldiam, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    


' Insert your code for output 4() create the verticle rectangular edge
    Dim xoff As Double
    xoff = NPD / 30
    
    pnts1(1).Set xoff, fd / 2 + FTK * 3, NPD / 2
    pnts1(2).Set xoff, VH - HWD * 1.2, NPD / 2
    pnts1(3).Set xoff, VH - HWD, NPD / 2.5
    pnts1(4).Set xoff, VH - HWD, NPD / 3
    pnts1(5).Set xoff, fd / 2 + FTK * 4, NPD / 3
    pnts1(6).Set xoff, fd / 2 + FTK * 4, -NPD / 3
    pnts1(7).Set xoff, VH - HWD, -NPD / 3
    pnts1(8).Set xoff, VH - HWD, -NPD / 2.5
    pnts1(9).Set xoff, VH - HWD * 1.2, -NPD / 2
    pnts1(10).Set xoff, fd / 2 + FTK * 3, -NPD / 2

    Set lines = New Collection
    For I = 1 To 9
        Set oLine = PlaceTrLine(pnts1(I), pnts1(I + 1))
        lines.Add oLine
    Next I
    Set oLine = PlaceTrLine(pnts1(10), pnts1(1))
    lines.Add oLine
    stPoint.Set pnts1(1).x, pnts1(1).y, pnts1(1).z
    Set objCStr = PlaceTrCString(stPoint, lines)
    Set oLine = Nothing
    
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
       
    Dim proVec As New AutoMath.DVector
    proVec.Set -1, 0, 0
    
    Dim objPro  As Object
    Set objPro = PlaceProjection(m_OutputColl, objCStr, proVec, xoff * 2, True)
    Set objCStr = Nothing
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPro
    Set objPro = Nothing
    xoff = NPD / 5
    
    
' Create the other support member
    
    pnts1(1).Set xoff, fd / 2 + FTK * 3, NPD / 2.7
    pnts1(2).Set xoff, VH - HWD, NPD / 2.7
    pnts1(3).Set xoff, VH - HWD, NPD / 3
    pnts1(4).Set xoff, fd / 2 + FTK * 4, NPD / 3
    pnts1(5).Set xoff, fd / 2 + FTK * 4, -NPD / 3
    pnts1(6).Set xoff, VH - HWD, -NPD / 3
    pnts1(7).Set xoff, VH - HWD, -NPD / 2.7
    pnts1(8).Set xoff, fd / 2 + FTK * 3, -NPD / 2.7

    Set lines = New Collection
    For I = 1 To 7
        Set oLine = PlaceTrLine(pnts1(I), pnts1(I + 1))
        lines.Add oLine
    Next I
    
    Set oLine = PlaceTrLine(pnts1(8), pnts1(1))
    lines.Add oLine
    
    stPoint.Set pnts1(1).x, pnts1(1).y, pnts1(1).z
    
    Set objCStr = PlaceTrCString(stPoint, lines)
    Set oLine = Nothing
    'Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
       
    proVec.Set -1, 0, 0

    Set objPro = PlaceProjection(m_OutputColl, objCStr, proVec, xoff * 2, True)
    Set objCStr = Nothing
        
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPro
    Set objPro = Nothing


'  create the transfercase
    Dim RPS(1 To 8) As New AutoMath.DPosition
    
    RPS(1).Set 0, VH - HWD * 0.2, FTK / 20
    RPS(2).Set 0, VH - HWD * 0.2, fd / 4
    RPS(3).Set 0, VH - HWD * 0.3, fd / 4
    RPS(4).Set 0, VH - HWD * 0.7, fd / 2
    RPS(5).Set 0, VH - HWD * 0.8, fd / 2
    RPS(6).Set 0, VH - HWD * 0.8, fd / 3
    RPS(7).Set 0, VH - HWD, fd / 3
    RPS(8).Set 0, VH - HWD, FTK / 20
    
    Set lines = New Collection
    For I = 1 To 7
        Set oLine = PlaceTrLine(RPS(I), RPS(I + 1))
        lines.Add oLine
    Next I
    
'''    Set oLine = PlaceTrLine(RPS(8), RPS(1))
'''    lines.Add oLine
    stPoint.Set RPS(1).x, RPS(1).y, RPS(1).z
    Set objCStr = PlaceTrCString(stPoint, lines)
    
    Set oLine = Nothing
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
    
    centPoint.Set 0, 0, 0
    axis.Set 0, 1, 0
    Set objRevolution = PlaceRevolution(m_OutputColl, objCStr, _
                            axis, centPoint, 2 * PI, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objRevolution = Nothing


' Spindle cone
    Dim objCone  As Object
    stPoint.Set 0, VH - HWD / 2, 0
    enPoint.Set Sin(HWA) * 1.1 * fd, VH - HWD / 2, Cos(HWA) * 1.1 * fd
    Set objCone = PlaceCone(m_OutputColl, stPoint, enPoint, HWD / 5, HWD / 15, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing

' handwheel spoke 1
    Dim BP As New DPosition
    Dim TP As New DPosition
    Dim objCyl  As Object
    TP.Set Sin(HWA) * fd, VH, Cos(HWA) * fd
    enPoint.Set Sin(HWA) * fd, VH - HWD, Cos(HWA) * fd
    BP.Set Sin(HWA) * fd, VH - HWD / 2, Cos(HWA) * fd
    Set objCyl = PlaceCylinder(m_OutputColl, TP, enPoint, HWD / 25, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing

' handwheel spoke 2
    stPoint.Set BP.x - Sin(HWA + 90 / RAD) * HWD / 2, BP.y, BP.z - Cos(HWA + 90 / RAD) * HWD / 2
    enPoint.Set BP.x + Sin(HWA + 90 / RAD) * HWD / 2, BP.y, BP.z + Cos(HWA + 90 / RAD) * HWD / 2
    Set objCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, HWD / 25, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing

' handwheel torus
    Dim objTorus    As Object
    Dim CP As New DPosition
    CP.Set BP.x, BP.y, BP.z
    axis.Set Sin(HWA), 0, Cos(HWA)
    Set objTorus = PlaceTorus(m_OutputColl, CP, axis, HWD / 2, HWD / 20)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTorus
    Set objTorus = Nothing

'handwheel knob
    enPoint.Set TP.x + Sin(HWA) * HWD / 5, TP.y, TP.z + Cos(HWA) * HWD / 5
    Set objCyl = PlaceCylinder(m_OutputColl, TP, enPoint, HWD / 25, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing




' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)
    oPlacePoint.Set -faceToFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set faceToFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set centPoint = Nothing
    Set axis = Nothing
    Set nvec = Nothing
    Set proVec = Nothing
    Set BP = Nothing
    Set TP = Nothing
    Set CP = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set gscadElem = Nothing
    Set oArc = Nothing
    Set Objo1 = Nothing
    Set Objo2 = Nothing
    Set Objo3 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

