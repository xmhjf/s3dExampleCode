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
'   Author:          SSP
'   Creation Date:  Monday, Oct 28 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IDL1D).
'   If this symbol is rotated about the Stem axis by 180 degrees ,Symbol IDR1A will be placed.
'   Physical aspect consist 1 Cylinder, 2 Cones ,1 stem ,1 box,1 dome cone,2 nozzles and 1 insulation aspect outputs.
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

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
 PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1       As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjConvergingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjValCylinder As Object
    Dim ObjStem As Object
    Dim ObjActuatorBox As Object
    Dim ObjActuatorCone As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double             'D2
    Dim parFacetoFace As Double                'D3
    Dim parCylHeight As Double                 'D4
    Dim parActuatorWidth As Double             'D5
    Dim parActuatorDiameter As Double          'D6
    Dim parActuatorHeight As Double            'D7
    Dim parActuatorHeight1 As Double           'D10
    Dim parActuatorHeight2 As Double           'D11
    Dim parActuatorLength As Double            'D12
    Dim parActuatorLength1 As Double           'D14
    Dim parActuatorWidth1 As Double            'D25
    Dim parActuatorWidth2 As Double            'D26
    Dim parActuatorLength2 As Double           'D28
    Dim parInsulationThickness As Double
    Dim parRotation As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorWidth = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight = arrayOfInputs(7)
    parActuatorHeight1 = arrayOfInputs(8)
    parActuatorHeight2 = arrayOfInputs(9)
    parActuatorLength = arrayOfInputs(10)
    parActuatorLength1 = arrayOfInputs(11)
    parActuatorWidth1 = arrayOfInputs(12)
    parActuatorWidth2 = arrayOfInputs(13)
    parActuatorLength2 = arrayOfInputs(14)
    parInsulationThickness = arrayOfInputs(15)
    parRotation = arrayOfInputs(16)
    
    iOutput = 0
' Insert your code for output 1(ConvergingCone)
    Dim pPos1   As New AutoMath.DPosition
    Dim pPos2   As New AutoMath.DPosition
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
    
    pPos1.Set -parFace1toCenter, 0, 0
    pPos2.Set 0, 0, 0
    
    Set ObjConvergingCone = PlaceCone(m_OutputColl, pPos1, pPos2, pipeDiam1 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergingCone
    Set ObjConvergingCone = Nothing
    
' Insert your code for output 2(DivergingCone)
    pPos1.Set parFacetoFace - parFace1toCenter, 0, 0
    pPos2.Set 0, 0, 0

    Set ObjDivergingCone = PlaceCone(m_OutputColl, pPos1, pPos2, pipeDiam2 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing
    
' Insert your code for output 3(ValCylinder)
    pPos1.Set 0, 0, 0
    pPos2.Set 0, -parCylHeight, 0
    
    Dim ValCylDiam As Double
' Assumption
' ValCylDiam = parFace1toCenter is taken
    ValCylDiam = parFace1toCenter
    
    Set ObjValCylinder = PlaceCylinder(m_OutputColl, pPos1, pPos2, ValCylDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCylinder
    Set ObjValCylinder = Nothing
    
' Insert your code for output 4(Stem)
    Dim StemDia As Double
    pPos1.Set 0, 0, 0
    pPos2.Set 0, parActuatorHeight - parActuatorHeight2, 0
' Assumption
' StemDia = 0.05 is taken
    StemDia = 0.05
    Set ObjStem = PlaceCylinder(m_OutputColl, pPos1, pPos2, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 5(ActuatorBox)
    Dim LineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Dim axisVector As New AutoMath.DVector
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
' Point 1 (Bottom Left)
    LineStrPts(0) = (parActuatorLength2 + parActuatorLength1)
    LineStrPts(1) = parActuatorHeight - parActuatorHeight2
    LineStrPts(2) = -parActuatorWidth2
' Point 2 (Bottom Right)
    LineStrPts(3) = (parActuatorLength2 + parActuatorLength1)
    LineStrPts(4) = parActuatorHeight - parActuatorHeight2
    LineStrPts(5) = parActuatorWidth1
' Point 3 (Top Right)
    LineStrPts(6) = (parActuatorLength2 + parActuatorLength1)
    LineStrPts(7) = parActuatorHeight + parActuatorHeight1
    LineStrPts(8) = parActuatorWidth1
' Point 4 (Top Left)
    LineStrPts(9) = (parActuatorLength2 + parActuatorLength1)
    LineStrPts(10) = parActuatorHeight + parActuatorHeight1
    LineStrPts(11) = -parActuatorWidth2
' Point 5 (Bottom Left)
    LineStrPts(12) = (parActuatorLength2 + parActuatorLength1)
    LineStrPts(13) = parActuatorHeight - parActuatorHeight2
    LineStrPts(14) = -parActuatorWidth2
    
    
    Dim ProjLength As Double
    ProjLength = parActuatorLength + parActuatorLength1
    Set oLString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)

    axisVector.Set -1, 0, 0
    Set ObjActuatorBox = PlaceProjection(m_OutputColl, oLString, axisVector, ProjLength, True)
    
    Set oLString = Nothing
' Rotation
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
' Axis of Rotation is Stem Axis
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjActuatorBox.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBox
    Set ObjActuatorBox = Nothing
    
' Insert your code for output 6(ActuatorCone)
' height of circular Dome is assumed as 0.2 times of the diameter - parActuatorDiameter
' Place Elliptical Arc on Right Side
    Dim oConeLine       As IngrGeom3D.Line3d
    Dim oDomeArc        As IngrGeom3D.Arc3d
    Dim oComplexStr     As IngrGeom3D.ComplexString3d
    Dim EleCollection   As Collection
    Dim Linepts(0 To 5) As Double
    Dim arcRadius       As Double
    Dim centerPoint   As New AutoMath.DPosition
    Dim startPoint   As New AutoMath.DPosition
    Dim endPoint   As New AutoMath.DPosition
    
        
    arcRadius = ((0.2 * parActuatorDiameter) ^ 2 + (parActuatorDiameter / 2) ^ 2) / _
                            (2 * 0.2 * parActuatorDiameter)
                            
    centerPoint.Set parActuatorLength2, parActuatorHeight, _
                    parActuatorWidth - arcRadius
    startPoint.Set parActuatorLength2, parActuatorHeight, parActuatorWidth
    endPoint.Set parActuatorLength2 + parActuatorDiameter / 2, parActuatorHeight, _
                    parActuatorWidth - 0.2 * parActuatorDiameter
    Set oDomeArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        centerPoint.x, centerPoint.y, centerPoint.z, _
                                                        startPoint.x, startPoint.y, startPoint.z, _
                                                        endPoint.x, endPoint.y, endPoint.z)
' Side of cone
    Linepts(0) = parActuatorLength2 + parActuatorDiameter / 2
    Linepts(1) = parActuatorHeight
    Linepts(2) = parActuatorWidth - 0.2 * parActuatorDiameter
    
    Linepts(3) = parActuatorLength2
    Linepts(4) = parActuatorHeight
    Linepts(5) = parActuatorWidth1
    Set oConeLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Set EleCollection = New Collection
    EleCollection.Add oDomeArc
    EleCollection.Add oConeLine
    
    Set oComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVector.Set 0, 0, 1
    centerPoint.Set parActuatorLength2, parActuatorHeight, parActuatorWidth1
' Revolve it about Z-Axis
    Set ObjActuatorCone = PlaceRevolution(m_OutputColl, oComplexStr, axisVector, centerPoint, PI * 2, True)
'Rotation
    ObjActuatorCone.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCone
    Set ObjActuatorCone = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oComplexStr
    Objcurves.Remove
    Set Objcurves = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To EleCollection.Count
        EleCollection.Remove 1
    Next iCount
    Set EleCollection = Nothing
    
    Set oDomeArc = Nothing
    Set oConeLine = Nothing
    Set oComplexStr = Nothing
    Set axisVector = Nothing
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
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set centerPoint = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set pPos1 = Nothing
    Set pPos2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


