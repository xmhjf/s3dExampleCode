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
'   Author:          MS
'   Creation Date:  Thursday, Oct 24 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=IVAQ)
'   The symbol consists of Nine outputs: Seven Physical outputs and Two Insulation Outputs
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjActuatorBody As Object
    Dim ObjPositionerBody As Object
   
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parPositionerHeight As Double
    Dim parPositionerOffset As Double
    Dim parPositionerOffset1 As Double
    Dim parPositionerOffset2 As Double
    Dim parValveOutertoCenter As Double
    Dim parPositionerLength As Double
    Dim parRotation As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parActuatorHeight = arrayOfInputs(4)    'D5
    parActuatorDiameter = arrayOfInputs(5)  'D6
    parPositionerHeight = arrayOfInputs(6)      'D10
    parPositionerOffset = arrayOfInputs(7)      'D11
    parPositionerOffset1 = arrayOfInputs(8)     'D12
    parPositionerOffset2 = arrayOfInputs(9)     'D14
    parValveOutertoCenter = arrayOfInputs(10)       'D19
    parPositionerLength = arrayOfInputs(11)     'D20
    parRotation = arrayOfInputs(12)
    parInsulationThickness = arrayOfInputs(13)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

 ' Insert your code for output 1(Body Cone1 of Valve)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 2(Body Cone2 of Valve)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set 0, -parFace2toCenter + flangeThick, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing

'    Insert your code for output 3(Valve Center Cylinder)
    Dim CylDia As Double
    Dim ValCenCylBottLength As Double
'   Valve Center Cylinder Diameter is taken to be 0.7 times the Nozzle2 diameter
    CylDia = 0.7 * pipeDiam
'   Valve Center Cylinder Bottom Length is taken to be 0.25 times the parFace2toCenter
    ValCenCylBottLength = parFace2toCenter * 0.25
    stPoint.Set 0, parValveOutertoCenter, 0
    enPoint.Set 0, -ValCenCylBottLength, 0
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)

 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
      
 ' Insert your code for output 4(Actuator Body)
     Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 8) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
'Create a lineString and Revolve it
 'Point 1
    LineStringPoints(0) = 0
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = parActuatorDiameter / 2
    LineStringPoints(4) = parValveOutertoCenter + parActuatorHeight
    LineStringPoints(5) = 0
'Point 3
    LineStringPoints(6) = 0
    LineStringPoints(7) = parValveOutertoCenter + parActuatorHeight
    LineStringPoints(8) = 0

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 3, LineStringPoints)
    axisVect.Set 0, 1, 0
    centerPoint.Set 0, (parValveOutertoCenter + parActuatorHeight) / 2, 0
    Set ObjActuatorBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    Set centerPoint = Nothing
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
    
 ' Insert your code for output 5(Positioner Body)
'Create a line string and project it
    Dim LineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d

'Point 1
    LineStrPts(0) = -(parPositionerLength - parPositionerOffset)
    LineStrPts(1) = parValveOutertoCenter + parPositionerOffset2 - parPositionerHeight
    LineStrPts(2) = parPositionerOffset1
'Point 2
    LineStrPts(3) = parPositionerOffset
    LineStrPts(4) = parValveOutertoCenter + parPositionerOffset2 - parPositionerHeight
    LineStrPts(5) = parPositionerOffset1
'Point 3
    LineStrPts(6) = parPositionerOffset
    LineStrPts(7) = parValveOutertoCenter + parPositionerOffset2
    LineStrPts(8) = parPositionerOffset1
'Point 4
    LineStrPts(9) = -(parPositionerLength - parPositionerOffset)
    LineStrPts(10) = parValveOutertoCenter + parPositionerOffset2
    LineStrPts(11) = parPositionerOffset1
'Point 5
    LineStrPts(12) = LineStrPts(0)
    LineStrPts(13) = LineStrPts(1)
    LineStrPts(14) = LineStrPts(2)
    
    Dim ProjLength As Double
    ProjLength = parPositionerOffset1 - (((parActuatorDiameter / 2) * (parValveOutertoCenter + parPositionerOffset2 - parPositionerHeight)) _
                            / (parValveOutertoCenter + parActuatorHeight))

    Set oLString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
     
    axisVect.Set 0, 0, -1
    Set ObjPositionerBody = PlaceProjection(m_OutputColl, oLString, axisVect, ProjLength, True)
    
'Transform the output
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjPositionerBody.Transform oTransMat
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPositionerBody
    Set ObjPositionerBody = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = oLString
    sp3dElem1.Remove
    Set oLString = Nothing
    Set oTransMat = Nothing
    Set oAboutVec = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset2 + depth2, 0
    oDir.Set 0, -1, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

