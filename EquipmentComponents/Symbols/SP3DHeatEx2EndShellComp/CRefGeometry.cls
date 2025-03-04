VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CRefGeometry.cls
'   Author:          ACM
'   Creation Date:  Tuesday, May 16, 2003
'   Description:
'   This is a Shell and Tube Exchanger shell Body.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

''  This is an Exchanger Shell Body Symbol.
'   This symbol is Fully parametric symbol, consists of 77 inputs and 10 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 8 are physical and 2 is ReferenceGeometry. It has Five nozzles and all
'   are fully parametric.
'   This is a Dow Equipment E-161 Exchanger Shell symbol.

Option Explicit

Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parVesselDiameter As Double
    Dim parVesselHeight As Double
    Dim parCPtoFace As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    Dim parNozzlePosition As Double
    Dim parNozzleOrientation As Double
    Dim parRadialPosition As Double
    Dim parOffset As Double
    Dim parNozzlePosition1 As Double
    Dim parNozzleOrientation1 As Double
    Dim parRadialPosition1 As Double
    Dim parOffset1 As Double
    Dim parNozzlePosition2 As Double
    Dim parNozzleOrientation2 As Double
    Dim parRadialPosition2 As Double
    Dim parOffset2 As Double
    Dim parNozzlePosition3 As Double
    Dim parNozzleOrientation3 As Double
    Dim parRadialPosition3 As Double
    Dim parOffset3 As Double
    Dim parNozzlePosition4 As Double
    Dim parNozzleOrientation4 As Double
    Dim parRadialPosition4 As Double
    Dim parOffset4 As Double
    
    Dim parNpd                     As Double
    Dim parEndPreparation          As Long
    Dim parScheduleThickness       As Long
    Dim parEndStandard             As Long
    Dim parPressureRating          As Long
    Dim parFlowDirection           As DistribFlow
    
    Dim parPortIndex1              As Long
    Dim parNpd1                    As Double
    Dim parEndPreparation1         As Long
    Dim parScheduleThickness1      As Long
    Dim parEndStandard1            As Long
    Dim parPressureRating1         As Long
    Dim parFlowDirection1          As DistribFlow
    
    Dim parPortIndex2              As Long
    Dim parNpd2                    As Double
    Dim parEndPreparation2         As Long
    Dim parScheduleThickness2      As Long
    Dim parEndStandard2            As Long
    Dim parPressureRating2         As Long
    Dim parFlowDirection2          As DistribFlow
    
    Dim parPortIndex3              As Long
    Dim parNpd3                    As Double
    Dim parEndPreparation3         As Long
    Dim parScheduleThickness3      As Long
    Dim parEndStandard3            As Long
    Dim parPressureRating3         As Long
    Dim parFlowDirection3          As DistribFlow
    
    Dim parPortIndex4              As Long
    Dim parNpd4                    As Double
    Dim parEndPreparation4         As Long
    Dim parScheduleThickness4      As Long
    Dim parEndStandard4            As Long
    Dim parPressureRating4         As Long
    Dim parFlowDirection4          As DistribFlow
    
    Dim parPortIndex5              As Long
    Dim parNpd5                    As Double
    Dim parEndPreparation5         As Long
    Dim parScheduleThickness5      As Long
    Dim parEndStandard5            As Long
    Dim parPressureRating5         As Long
    Dim parFlowDirection5          As DistribFlow
    
    Dim parId1                     As String
    Dim parId2                     As String
    Dim parId3                     As String
    Dim parId4                     As String
    Dim parId5                     As String
    Dim parNPDUnitType             As String
    Dim parNpdUnitType1            As String
    Dim parNpdUnitType2            As String
    Dim parNpdUnitType3            As String
    Dim parNpdUnitType4            As String
    Dim parNpdUnitType5            As String
        
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselHeight = arrayOfInputs(3)
    parCPtoFace = arrayOfInputs(4)
    parFaceDiameter = arrayOfInputs(5)
    parFaceThickness = arrayOfInputs(6)
    parNozzlePosition = arrayOfInputs(7)
    parNozzleOrientation = arrayOfInputs(8)
    parRadialPosition = arrayOfInputs(9)
    parOffset = arrayOfInputs(10)
    parNozzlePosition1 = arrayOfInputs(11)
    parNozzleOrientation1 = arrayOfInputs(12)
    parRadialPosition1 = arrayOfInputs(13)
    parOffset1 = arrayOfInputs(14)
    parNozzlePosition2 = arrayOfInputs(15)
    parNozzleOrientation2 = arrayOfInputs(16)
    parRadialPosition2 = arrayOfInputs(17)
    parOffset2 = arrayOfInputs(18)
    parNozzlePosition3 = arrayOfInputs(19)
    parNozzleOrientation3 = arrayOfInputs(20)
    parRadialPosition3 = arrayOfInputs(21)
    parOffset3 = arrayOfInputs(22)
    parNozzlePosition4 = arrayOfInputs(23)
    parNozzleOrientation4 = arrayOfInputs(24)
    parRadialPosition4 = arrayOfInputs(25)
    parOffset4 = arrayOfInputs(26)

    parNpd = arrayOfInputs(27)
    parEndPreparation = arrayOfInputs(28)
    parScheduleThickness = arrayOfInputs(29)
    parEndStandard = arrayOfInputs(30)
    parPressureRating = arrayOfInputs(31)
    parFlowDirection = arrayOfInputs(32)
    parPortIndex1 = arrayOfInputs(33)
    parNpd1 = arrayOfInputs(34)
    parEndPreparation1 = arrayOfInputs(35)
    parScheduleThickness1 = arrayOfInputs(36)
    parEndStandard1 = arrayOfInputs(37)
    parPressureRating1 = arrayOfInputs(38)
    parFlowDirection1 = arrayOfInputs(39)
    parPortIndex2 = arrayOfInputs(40)
    parNpd2 = arrayOfInputs(41)
    parEndPreparation2 = arrayOfInputs(42)
    parScheduleThickness2 = arrayOfInputs(43)
    parEndStandard2 = arrayOfInputs(44)
    parPressureRating2 = arrayOfInputs(45)
    parFlowDirection2 = arrayOfInputs(46)
    parPortIndex3 = arrayOfInputs(47)
    parNpd3 = arrayOfInputs(48)
    parEndPreparation3 = arrayOfInputs(49)
    parScheduleThickness3 = arrayOfInputs(50)
    parEndStandard3 = arrayOfInputs(51)
    parPressureRating3 = arrayOfInputs(52)
    parFlowDirection3 = arrayOfInputs(53)
    parPortIndex4 = arrayOfInputs(54)
    parNpd4 = arrayOfInputs(55)
    parEndPreparation4 = arrayOfInputs(56)
    parScheduleThickness4 = arrayOfInputs(57)
    parEndStandard4 = arrayOfInputs(58)
    parPressureRating4 = arrayOfInputs(59)
    parFlowDirection4 = arrayOfInputs(60)
    parPortIndex5 = arrayOfInputs(61)
    parNpd5 = arrayOfInputs(62)
    parEndPreparation5 = arrayOfInputs(63)
    parScheduleThickness5 = arrayOfInputs(64)
    parEndStandard5 = arrayOfInputs(65)
    parPressureRating5 = arrayOfInputs(66)
    parFlowDirection5 = arrayOfInputs(67)
    
    parId1 = arrayOfInputs(68)
    parId2 = arrayOfInputs(69)
    parId3 = arrayOfInputs(70)
    parId4 = arrayOfInputs(71)
    parId5 = arrayOfInputs(72)
    parNPDUnitType = arrayOfInputs(73)
    parNpdUnitType1 = arrayOfInputs(74)
    parNpdUnitType2 = arrayOfInputs(75)
    parNpdUnitType3 = arrayOfInputs(76)
    parNpdUnitType4 = arrayOfInputs(77)
    parNpdUnitType5 = arrayOfInputs(78)

    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

' Insert your code for output 9( Point at origin )
    Dim ObjExchangerPoint As IngrGeom3D.Point3d
    
    Set ObjExchangerPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerPoint
    Set ObjExchangerPoint = Nothing
    
'   Create DefaultSurface - for output 10(a Circular shape is chosen)
    Dim ObjDefaultSurface As Object
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, 1, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
