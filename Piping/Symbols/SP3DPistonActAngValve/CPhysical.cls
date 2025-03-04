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
'   Creation Date:  Wednesday, Oct 16 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=ISSPA
'   The symbol consists of 2 Cones, 1 Revolution object as Actuator cum Valve Cylinder, and 2 Nozzles as
'   part of physical aspect. Insulation aspect consist of 2 Cylinders for valve body
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
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjActuatorCylinder As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl  As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)          'D2
    parFace2toCenter = arrayOfInputs(3)          'D3
    parActuatorHeight = arrayOfInputs(4)           'D5
    parActuatorDiameter = arrayOfInputs(5)           'D6
    parInstrumentHeight = arrayOfInputs(6)           'D19
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
' Origin is assumed to be at the center of the valveBody along the flow axis
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
    
 ' Insert your code for output 1(Actuator Cylinder)
    Dim actuatorDiameter1 As Double
    Dim ValCenCylBottLength As Double
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 17) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    'Valve Center Cylinder Diameter is taken to be 0.7 times the Nozzle2 diameter
    actuatorDiameter1 = 0.7 * pipeDiam
    'Valve Center Cylinder Bottom Length is taken to be 0.25 times the parFace2toCenter
    ValCenCylBottLength = parFace2toCenter * 0.25

'Create a lineString and Revolve it
 'Point 1
    LineStringPoints(0) = 0
    LineStringPoints(1) = -ValCenCylBottLength
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = actuatorDiameter1 / 2
    LineStringPoints(4) = -ValCenCylBottLength
    LineStringPoints(5) = 0
'Point 3
    LineStringPoints(6) = actuatorDiameter1 / 2
    LineStringPoints(7) = parInstrumentHeight
    LineStringPoints(8) = 0
'Point 4
    LineStringPoints(9) = parActuatorDiameter / 2
    LineStringPoints(10) = parInstrumentHeight
    LineStringPoints(11) = 0
'Point 5
    LineStringPoints(12) = parActuatorDiameter / 2
    LineStringPoints(13) = parActuatorHeight + parInstrumentHeight
    LineStringPoints(14) = 0
'Point 6
    LineStringPoints(15) = 0
    LineStringPoints(16) = parActuatorHeight + parInstrumentHeight
    LineStringPoints(17) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStringPoints)
    axisVect.Set 0, 1, 0
    centerPoint.Set 0, (parActuatorHeight + parInstrumentHeight) / 2, 0
    Set ObjActuatorCylinder = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCylinder
    Set ObjActuatorCylinder = Nothing
    Set centerPoint = Nothing
    Set axisVect = Nothing
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1 (Nozzle1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2 (Nozzle2)
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset2 + depth2, 0
    oDir.Set 0, -1, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
