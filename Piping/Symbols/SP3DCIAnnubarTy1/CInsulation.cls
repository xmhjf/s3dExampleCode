VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  13 Monday, January2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IA1,
'   PDMS Category Name=IFAI). '   Physical aspect consist a line string revolution about X-axis covering two
'   Physical aspect consist a line string revolution about X-axis covering two
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
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata As IJDCodeListMetaData
Private Sub Class_Initialize()

    PI = 4 * Atn(1)

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt As PartFacelets.IJDPart
    Dim ObjAnubarStemIns As Object
    Dim ObjAnuTeeBodyIns As Object
    Dim ObjBotElbowIns As Object
    Dim ObjHoriBotBodyIns As Object
    Dim ObjTopElbowIns As Object
    Dim ObjHoriTopBodyIns As Object
    
    Dim pipeDia(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    Dim flangeDia(1 To 1)  As Double
    Dim sptOffset(1 To 1) As Double
    Dim depth(1 To 1) As Double
    Dim iOutput As Double
    
    Dim parInstrumentLength As Double
    Dim parFacetoEnd As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInsulationThickness As Double
    'Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
    
    Dim parOperationHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)              'D2
    parFacetoEnd = arrayOfInputs(3)                             'D3
    parInstrumentLength1 = arrayOfInputs(4)          'D4
    parInstrumentDiameter = arrayOfInputs(5)          'D5
    parInstrumentHeight = arrayOfInputs(6)              'D6
    parInstrumentHeight1 = arrayOfInputs(7)             'D7
    parInsulationThickness = arrayOfInputs(8)
    parNPD = arrayOfInputs(9)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(10)
    parScheduleThickness = arrayOfInputs(11)
    parEndStandard = arrayOfInputs(12)
    parPressureRating = arrayOfInputs(13)
    parFlowDirection = arrayOfInputs(14)
    parPortIndex1 = arrayOfInputs(15)
    parNpd1 = arrayOfInputs(16)
    parEndPreparation1 = arrayOfInputs(17)
    parScheduleThickness1 = arrayOfInputs(18)
    parEndStandard1 = arrayOfInputs(19)
    parPressureRating1 = arrayOfInputs(20)
    parFlowDirection1 = arrayOfInputs(21)
    parOperationHeight = arrayOfInputs(22)          'D8
    
    parId1 = arrayOfInputs(23)
    parNPDUnitType = arrayOfInputs(24)
    parNPDUnitType1 = arrayOfInputs(25)
    
    iOutput = 0
    
' Insert your code for output 5(Valve Body Insulation)
' m_oCodeListMetadata not initialised as it's a dummy parameter for the function
     CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDia, flangeThick, _
                                            flangeDia, sptOffset, depth

' Insert your code for output 7(Annubar Stem Insulation)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
    
    stPoint.Set 0, 0, 0
    enPoint.Set parFacetoEnd - parInstrumentLength1, 0, 0
    InsulationDia = parInstrumentDiameter + 2 * parInsulationThickness
    If flangeDia(parPortIndex1) > parInstrumentDiameter Then
        InsulationDia = flangeDia(parPortIndex1) + 2 * parInsulationThickness
    End If
    Set ObjAnubarStemIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAnubarStemIns
    Set ObjAnubarStemIns = Nothing

'   Insert your code for output 8(Annu Bar Tee Body Insulation)
'   Create a lineString and Revolve it
    Dim minCylDia As Double
    
    minCylDia = 0.015    'As per PDS Eden Code
    stPoint.Set parFacetoEnd - parInstrumentLength1, _
                        parInstrumentHeight - 1.5 * minCylDia, 0
    enPoint.Set parFacetoEnd - parInstrumentLength1, _
                        -parInstrumentHeight1 + 1.5 * minCylDia, 0
    InsulationDia = parInstrumentDiameter + 2 * parInsulationThickness
    If minCylDia > parInstrumentDiameter Then InsulationDia = minCylDia + 2 * parInsulationThickness
    Set ObjAnuTeeBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAnuTeeBodyIns
    Set ObjAnuTeeBodyIns = Nothing
    
 ' Insert your code for output 9(Bottom Elbow)
    Dim majRadius As Double
    Dim minorRad As Double
    Dim oElbowCir As IngrGeom3D.Circle3d
    Dim cirCenterPt As AutoMath.DPosition
    Dim centerPoint As AutoMath.DPosition
    Dim axisVect As AutoMath.DVector
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim InsulationRadius As Double
    
    Set centerPoint = New AutoMath.DPosition
    Set cirCenterPt = New AutoMath.DPosition
    Set axisVect = New AutoMath.DVector
        
    cirCenterPt.Set parFacetoEnd - parInstrumentLength1, _
                            -parInstrumentHeight1 + minCylDia * 1.5, 0
    majRadius = 1.5 * minCylDia     '1.5 times of torus mini Radius as per Eden code
    minorRad = minCylDia / 2
    centerPoint.Set parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5, _
                                -parInstrumentHeight1 + minCylDia * 1.5, 0
    axisVect.Set 0, 0, 1
    InsulationRadius = (minorRad + parInsulationThickness)
    If InsulationRadius > majRadius Then InsulationRadius = majRadius
    Set oElbowCir = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                            cirCenterPt.x, cirCenterPt.y, cirCenterPt.z, _
                                                                            0, 1, 0, InsulationRadius)
    Set ObjBotElbowIns = PlaceRevolution(m_OutputColl, oElbowCir, axisVect, centerPoint, PI / 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotElbowIns
    Set ObjBotElbowIns = Nothing
    Set oElbowCir = Nothing
    
 ' Insert your code for output 10(Bottom Horizontal Cylinder and Cones)
    stPoint.Set parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5, _
                        -parInstrumentHeight1, 0
    enPoint.Set parFacetoEnd + parInsulationThickness, _
                        -parInstrumentHeight1, 0
    InsulationDia = minCylDia + 2 * parInsulationThickness
    Set ObjHoriBotBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriBotBodyIns
    Set ObjHoriBotBodyIns = Nothing

' Insert your code for output 11(Top Elbow)
    majRadius = 1.5 * minCylDia     '1.5 times of torus mini Radius as per Eden code
    minorRad = minCylDia / 2
    cirCenterPt.Set parFacetoEnd - parInstrumentLength1, _
                                parInstrumentHeight - minCylDia * 1.5, 0
    centerPoint.Set parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5, _
                                parInstrumentHeight + minCylDia * 1.5, 0
    axisVect.Set 0, 0, 1
    InsulationRadius = (minorRad + parInsulationThickness)
    If InsulationRadius > majRadius Then InsulationRadius = majRadius
    Set oElbowCir = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                            cirCenterPt.x, cirCenterPt.y, cirCenterPt.z, _
                                                                            0, 1, 0, InsulationRadius)
    Set ObjTopElbowIns = PlaceRevolution(m_OutputColl, oElbowCir, axisVect, centerPoint, PI / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopElbowIns
    Set ObjTopElbowIns = Nothing
    Set oElbowCir = Nothing
    Set cirCenterPt = Nothing
    Set axisVect = Nothing
    Set centerPoint = Nothing
    Set geomFactory = Nothing
    
' Insert your code for output 12(Top Horizontal Cylinder and Cones)
    stPoint.Set parFacetoEnd - parInstrumentLength1 + minCylDia * 1.5, _
                        parInstrumentHeight, 0
    enPoint.Set parFacetoEnd + parInsulationThickness, _
                        parInstrumentHeight, 0
    InsulationDia = minCylDia + 2 * parInsulationThickness
    Set ObjHoriTopBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriTopBodyIns
    Set ObjHoriTopBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
