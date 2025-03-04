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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Wednusday 14, Feb 2007
'   Description:
'   This implements physical aspect of the symbol created for Sump Strainer as per CR-113397
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjCylinder As Object
    Dim objNut As Object
    
    Dim parFacetoEnd As Double
    Dim parDiameter As Double
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dNutHeight As Double
    Dim PI As Double
    PI = 4 * Atn(1)
    
    Dim InchToMtr As Double
    InchToMtr = 25.4 / 1000
    
    Dim half As Double
    Dim two As Double
    half = 0.5 * InchToMtr
    two = 2 * InchToMtr

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoEnd = arrayOfInputs(2)
    parDiameter = arrayOfInputs(3)
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
'   Place Output 1 (Cylinder for body)
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
'Assumptions for Nut Height
    If CmpDblLessThanOrEqualTo(pipeDiam, half) Then
             dNutHeight = pipeDiam * 0.9
    ElseIf CmpDblGreaterthan(pipeDiam, half) And CmpDblLessThanOrEqualTo(pipeDiam, two) Then
              dNutHeight = pipeDiam * 0.75
    ElseIf CmpDblGreaterthan(pipeDiam, two) Then
                dNutHeight = pipeDiam * 0.6
    End If
    
    oStPoint.Set dNutHeight, 0, 0
    oEnPoint.Set parFacetoEnd, 0, 0
    
    Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing

'   Place Output 2(Nut)
    
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objHexagon         As IngrGeom3D.ComplexString3d
    
    Dim dAngle           As Double
    Dim dHexaZ            As Double
    Dim dHexaY            As Double
    Dim dNutWidth As Double
    
    Dim arrHexagon(21)     As Double
    dAngle = PI / 3
    dHexaY = 0.55 * pipeDiam
    dHexaZ = dHexaY / Tan(dAngle)
    dNutWidth = dHexaY / Sin(dAngle)
    
' Point 1
    arrHexagon(0) = 0
    arrHexagon(1) = 0
    arrHexagon(2) = dNutWidth
' Point 2
    arrHexagon(3) = 0
    arrHexagon(4) = -dHexaY
    arrHexagon(5) = dHexaZ
' Point 3
    arrHexagon(6) = 0
    arrHexagon(7) = -dHexaY
    arrHexagon(8) = -dHexaZ
' Point 4
    arrHexagon(9) = 0
    arrHexagon(10) = 0
    arrHexagon(11) = -dNutWidth
' Point 5
    arrHexagon(12) = 0
    arrHexagon(13) = dHexaY
    arrHexagon(14) = -dHexaZ
' Point 6
    arrHexagon(15) = 0
    arrHexagon(16) = dHexaY
    arrHexagon(17) = dHexaZ

    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrHexagon(0), arrHexagon(1), arrHexagon(2), arrHexagon(3), arrHexagon(4), arrHexagon(5))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrHexagon(3), arrHexagon(4), arrHexagon(5), arrHexagon(6), arrHexagon(7), arrHexagon(8))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrHexagon(6), arrHexagon(7), arrHexagon(8), arrHexagon(9), arrHexagon(10), arrHexagon(11))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrHexagon(9), arrHexagon(10), arrHexagon(11), arrHexagon(12), arrHexagon(13), arrHexagon(14))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrHexagon(12), arrHexagon(13), arrHexagon(14), arrHexagon(15), arrHexagon(16), arrHexagon(17))
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrHexagon(15), arrHexagon(16), arrHexagon(17), arrHexagon(0), arrHexagon(1), arrHexagon(2))
    lines.Add oLine
    
    oStPoint.Set arrHexagon(0), arrHexagon(1), arrHexagon(2)
    Set objHexagon = PlaceTrCString(oStPoint, lines)
    
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing

    Set objNut = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        objHexagon, 1, 0, 0, dNutHeight, True)
    Set objHexagon = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNut
    Set objNut = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double
    
    oPlacePoint.Set -sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD


End Sub

