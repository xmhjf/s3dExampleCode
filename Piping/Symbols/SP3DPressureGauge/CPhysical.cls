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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         dkl
'   Creation Date:  Thursday, Jul 5 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
''     Pressure gauge (Source: 750 Series at www.trerice.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI       As Double
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize"
    On Error GoTo Errx

        PI = 4 * Atn(1)
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt         As PartFacelets.IJDPart
    Dim flangeThick       As Double
    Dim sptOffset         As Double
    Dim flangeDiam        As Double
    Dim pipeDiam          As Double
    Dim Depth             As Double
   
    Dim iOutput    As Double
            
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentDiameter = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    parInstrumentWidth = arrayOfInputs(4)
    parInstrumentWidth1 = arrayOfInputs(5)

    iOutput = 0
 
' output 1(Threaded insertion)
    Dim ObjThreadedInsertion As Object
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Dim dThreadedInsLength As Double
    Dim dNutHeight         As Double
    Dim k As Double
    Dim dHexDia            As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, Depth
    
    k = 0.85    ' Assumption : Nut Height = K * Pipe Dia
    Const F = 0.00127  ' Gauge radius = parInstrumentDiameter - F. Assumption : F = 0.05"
    dHexDia = 1.5 * pipeDiam  'Assumption:Distance between two faces of hexagon(dHexDia) = 1.5 * pipeDiam
    
    dNutHeight = k * pipeDiam
    dThreadedInsLength = parInstrumentHeight - (parInstrumentDiameter / 2) - dNutHeight
    If CmpDblLessThanOrEqualTo(dThreadedInsLength, 0) Then
        k = 0.4
        dNutHeight = k * pipeDiam
        dThreadedInsLength = parInstrumentHeight - (parInstrumentDiameter / 2) - dNutHeight
     End If

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set 0, 0, 0
'   Note: This cylinder would show the portion below and above the Hexagonal Nut
    enPoint.Set parInstrumentHeight, 0, 0
    Set ObjThreadedInsertion = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)

' Set the output 1
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjThreadedInsertion
    Set ObjThreadedInsertion = Nothing
    
' Output 2 (The heaxagonal nut)
    Dim ObjNut   As Object
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim iCount            As Integer
    Dim dAngle             As Double

    Dim axisVect   As AutoMath.DVector
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d

    dAngle = PI / 3
    stPoint.Set dThreadedInsLength, 0, 0
'   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = stPoint.x
        LineStrPoints(3 * iCount - 2) = stPoint.y + (dHexDia / 2) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints(3 * iCount - 1) = stPoint.z + (dHexDia / 2) * Sin(PI / 6 + iCount * dAngle)
    Next iCount
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set axisVect = New AutoMath.DVector
    axisVect.Set 1, 0, 0
    Set ObjNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)

' Set the output 2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut
    Set ObjNut = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing
    
' output 3 (Gauge cylinder)
    Dim ObjGauge As Object

    stPoint.Set parInstrumentHeight, 0, -parInstrumentWidth1
    enPoint.Set parInstrumentHeight, 0, pipeDiam / 2
    Set ObjGauge = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter - F, True)

' Set the output 3
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjGauge
    Set ObjGauge = Nothing
    
' output 4 (Dial cylinder)
    Dim ObjDial As Object

    stPoint.Set parInstrumentHeight, 0, pipeDiam / 2
    enPoint.Set parInstrumentHeight, 0, parInstrumentWidth - parInstrumentWidth1
    Set ObjDial = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)

' Set the output 4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDial
    Set ObjDial = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set 0, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
