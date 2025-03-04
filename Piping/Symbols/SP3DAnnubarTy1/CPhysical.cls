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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IA1,
'   PDMS Category Name=IFAI). Physical aspect consist a line string revolution about X-axis covering two
'   cones  and a single cylinder, another line string revolution which generates Two cones and a Single Cylinder,
'   a bottom Elbow, a line string revolution   about X-axis generating a cylinder and 2 cones Two cones at Bottom, a top
'   Elbow , a line string revolution   about X-axis generating a cylinder and 2 cones Two cones at Top and 1 Nozzle.
'   Insulation consist of four Cylinders and 2 Elbows.
'
'   Note: This symbol expects the user to place an Olet component on the pipe and this (Annubar) should be placed
'   at the connect point2  of the Olet.The parameters values are to be taken from connect point2 of the Olet.
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
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim ObjAnubarStem As Object
    Dim ObjAnuTeeBody As Object
    Dim ObjBotElbow As Object
    Dim ObjHoriBotBody As Object
    Dim ObjTopElbow As Object
    Dim ObjHoriTopBody As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    
    Dim parInstrumentLength As Double
    Dim parFacetoEnd As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)              'D2
    parFacetoEnd = arrayOfInputs(3)                             'D3
    parInstrumentLength1 = arrayOfInputs(4)          'D4
    parInstrumentDiameter = arrayOfInputs(5)          'D5
    parInstrumentHeight = arrayOfInputs(6)              'D6
    parInstrumentHeight1 = arrayOfInputs(7)             'D7
    parInsulationThickness = arrayOfInputs(8)         'D8
    
    iOutput = 0
    
    'Origin of the Symbol is assumed at Nozzle-1 position
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
 ' Insert your code for output 1(Annubar Stem)
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 17) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
' Code for Cone 1
    stPoint.Set flangeThick, 0, 0
    enPoint.Set (parInstrumentLength - flangeThick) / 2, 0, 0#
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing

'Code for Cone 2
    stPoint.Set (parInstrumentLength - flangeThick) / 2, 0, 0
    enPoint.Set parInstrumentLength, 0, 0#
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, enPoint, stPoint, parInstrumentDiameter / 2, 0#, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
'Code for Cylinder
    stPoint.Set parInstrumentLength - flangeThick, 0, 0
    enPoint.Set parFacetoEnd - parInstrumentLength1, 0, 0#
    Set ObjAnubarStem = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAnubarStem
    Set ObjAnubarStem = Nothing

'   Insert your code for output 2(Annu Bar Tee Body)
'   Create a lineString and Revolve it
    Dim minCylDia As Double
    Dim teeLength As Double
    minCylDia = 0.015    'As per PDS Eden Code
    teeLength = parInstrumentDiameter * 1.25
 'Point 1
    LineStringPoints(0) = parFacetoEnd - parInstrumentLength1
    LineStringPoints(1) = parInstrumentHeight - minCylDia * 1.5 'Elbow Radius is 1.5*Minimum Cylin dia
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = parFacetoEnd - parInstrumentLength1 + minCylDia / 2
    LineStringPoints(4) = parInstrumentHeight - minCylDia * 1.5
    LineStringPoints(5) = 0
'Point 3
    LineStringPoints(6) = parFacetoEnd - parInstrumentLength1 _
                                        + parInstrumentDiameter / 2
    LineStringPoints(7) = teeLength / 2
    LineStringPoints(8) = 0
'Point 4
    LineStringPoints(9) = parFacetoEnd - parInstrumentLength1 _
                                        + parInstrumentDiameter / 2
    LineStringPoints(10) = -teeLength / 2
    LineStringPoints(11) = 0
'Point 5
    LineStringPoints(12) = parFacetoEnd - parInstrumentLength1 + minCylDia / 2
    LineStringPoints(13) = -parInstrumentHeight1 + minCylDia * 1.5
    LineStringPoints(14) = 0
'Point 6
    LineStringPoints(15) = parFacetoEnd - parInstrumentLength1
    LineStringPoints(16) = -parInstrumentHeight1 + minCylDia * 1.5
    LineStringPoints(17) = 0

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStringPoints)
    axisVect.Set 0, 1, 0
    centerPoint.Set parFacetoEnd - parInstrumentLength1, 0, 0
    Set ObjAnuTeeBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAnuTeeBody
    Set ObjAnuTeeBody = Nothing
    
 ' Insert your code for output 3(Bottom Elbow)
    Dim majRadius As Double
    Dim minorRad As Double
    Dim oElbowCir As IngrGeom3D.Circle3d
    Dim cirCenterPt As AutoMath.DPosition
    Set cirCenterPt = New AutoMath.DPosition
    
    cirCenterPt.Set parFacetoEnd - parInstrumentLength1, _
                            -parInstrumentHeight1 + minCylDia * 1.5, 0
    majRadius = 1.5 * minCylDia     '1.5 times of torus mini Radius as per Eden code
    minorRad = minCylDia / 2
    centerPoint.Set parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5, _
                                -parInstrumentHeight1 + minCylDia * 1.5, 0
    axisVect.Set 0, 0, 1
    Set oElbowCir = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                            cirCenterPt.x, cirCenterPt.y, cirCenterPt.z, _
                                                                            0, 1, 0, minorRad)
    Set ObjBotElbow = PlaceRevolution(m_OutputColl, oElbowCir, axisVect, centerPoint, PI / 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotElbow
    Set ObjBotElbow = Nothing
    Set oElbowCir = Nothing
    
 ' Insert your code for output 4(Bottom Horizontal Cylinder and Cones)
'Point 1
    LineStringPoints(0) = parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5
    LineStringPoints(1) = -parInstrumentHeight1
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5
    LineStringPoints(4) = -parInstrumentHeight1 + minCylDia / 2
    LineStringPoints(5) = 0
'Point 3
    LineStringPoints(6) = parFacetoEnd - parInstrumentLength1 * 0.25 * 2
    LineStringPoints(7) = -parInstrumentHeight1 + minCylDia / 2
    LineStringPoints(8) = 0
'Point 4
    LineStringPoints(9) = parFacetoEnd - parInstrumentLength1 * 0.25
    LineStringPoints(10) = -parInstrumentHeight1
    LineStringPoints(11) = 0.00001  'we need this value otherwise code will give assertion failure
'Point 5
    LineStringPoints(12) = parFacetoEnd
    LineStringPoints(13) = -parInstrumentHeight1 + minCylDia / 2
    LineStringPoints(14) = 0
'Point 6
    LineStringPoints(15) = parFacetoEnd
    LineStringPoints(16) = -parInstrumentHeight1
    LineStringPoints(17) = 0
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set parFacetoEnd - parInstrumentLength1 * 0.25, -parInstrumentHeight1, 0

    Set ObjHoriBotBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriBotBody
    Set ObjHoriBotBody = Nothing

' Insert your code for output 5(Top Elbow)
    cirCenterPt.Set parFacetoEnd - parInstrumentLength1, _
                                parInstrumentHeight - minCylDia * 1.5, 0
    centerPoint.Set parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5, _
                                parInstrumentHeight - minCylDia * 1.5, 0
    axisVect.Set 0, 0, -1
    Set oElbowCir = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                            cirCenterPt.x, cirCenterPt.y, cirCenterPt.z, _
                                                                            0, -1, 0, minorRad)

    Set ObjTopElbow = PlaceRevolution(m_OutputColl, oElbowCir, axisVect, centerPoint, PI / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopElbow
    Set ObjTopElbow = Nothing
    Set oElbowCir = Nothing
    Set cirCenterPt = Nothing
    
' Insert your code for output 6(Top Horizontal Cylinder and Cones)
'Point 1
    LineStringPoints(0) = parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5
    LineStringPoints(1) = parInstrumentHeight
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5
    LineStringPoints(4) = parInstrumentHeight + minCylDia / 2
    LineStringPoints(5) = 0
'Point 3
    LineStringPoints(6) = parFacetoEnd - parInstrumentLength1 * 0.25 * 2
    LineStringPoints(7) = parInstrumentHeight + minCylDia / 2
    LineStringPoints(8) = 0
'Point 4
    LineStringPoints(9) = parFacetoEnd - parInstrumentLength1 * 0.25
    LineStringPoints(10) = parInstrumentHeight
    LineStringPoints(11) = 0.00001
'Point 5
    LineStringPoints(12) = parFacetoEnd
    LineStringPoints(13) = parInstrumentHeight + minCylDia / 2
    LineStringPoints(14) = 0
'Point 6
    LineStringPoints(15) = parFacetoEnd
    LineStringPoints(16) = parInstrumentHeight
    LineStringPoints(17) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set parFacetoEnd - parInstrumentLength1 * 0.25, parInstrumentHeight, 0

    Set ObjHoriTopBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriTopBody
    Set ObjHoriTopBody = Nothing
    Set axisVect = Nothing
    Set centerPoint = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -sptOffset + depth, 0, 0
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
    Resume Next
    
End Sub
