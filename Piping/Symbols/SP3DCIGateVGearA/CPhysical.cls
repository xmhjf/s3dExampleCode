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
'   Creation Date:  Sunday, Dec 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.Feb.2007    RRK           TR-113129 Changes made to include Insulation Aspect
'   07-Mar-07      RRK           TR-116058: Added 'If' condition to suppport new optional input without failing the prior to V7 instances
'  1.NOV.2007      RRK           CR-123952 Updated the symbol to use CreateRetrieveDynamicNozzle function in FullyParametricFun.bas
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI                      As Double
Private RAD                     As Double
Private gscadElem               As IJDObject
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI
End Sub

Private Sub Class_Terminate()

    Set m_oCodeListMetadata = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"

    On Error GoTo ErrorLabel
    
    Dim i                       As Integer
    Dim oPartFclt               As PartFacelets.IJDPart
    
    Dim dBodyFlangeThick As Double
    Dim dBodyFlangeDiam As Double
    
    Dim pipeDiam(1 To 2) As Double
    Dim CptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double

    Dim iOutput                 As Double
    
    Dim parValveWidth           As Double
    Dim parValveHeight          As Double
    Dim parHandwheelDiameter    As Double
    Dim parHandwheelAngle       As Double
    
    Dim Npd                     As Double
    Dim EndPreparation          As Long
    Dim ScheduleThickness       As Long
    Dim EndStandard             As Long
    Dim PressureRating          As Long
    Dim FlowDirection           As DistribFlow
    
    Dim PortIndex1              As Long
    Dim Npd1                    As Double
    Dim EndPreparation1         As Long
    Dim ScheduleThickness1      As Long
    Dim EndStandard1            As Long
    Dim PressureRating1         As Long
    Dim FlowDirection1          As DistribFlow
'Comment to make test with text inputs
'    Dim Id1                     As Long
'-------------------------------------
    Dim PortIndex2              As Long
    Dim Npd2                    As Double
    Dim EndPreparation2         As Long
    Dim ScheduleThickness2      As Long
    Dim EndStandard2            As Long
    Dim PressureRating2         As Long
    Dim FlowDirection2          As DistribFlow
    Dim parInsulationThickness As Double
    
    Dim Id1                     As String
    Dim Id2                     As String
''---------------------------------------
    Dim NpdUnitType             As String
    Dim NpdUnitType1            As String
    Dim NpdUnitType2            As String
    '
    Dim TerminationClass1       As Long
    Dim TerminationSubClass1    As Long
    Dim SchedulePractice1       As Long
    Dim EndPractice1            As Long
    Dim RatingPractice1         As Long
    Dim TerminationClass2       As Long
    Dim TerminationSubClass2    As Long
    Dim SchedulePractice2       As Long
    Dim EndPractice2            As Long
    Dim RatingPractice2         As Long
    
    'Numeric Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveWidth = arrayOfInputs(2)
    parValveHeight = arrayOfInputs(3)
    parHandwheelDiameter = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    Npd = arrayOfInputs(6)
    EndPreparation = arrayOfInputs(7)
    ScheduleThickness = arrayOfInputs(8)
    EndStandard = arrayOfInputs(9)
    PressureRating = arrayOfInputs(10)
    FlowDirection = arrayOfInputs(11)
    PortIndex1 = arrayOfInputs(12)
    Npd1 = arrayOfInputs(13)
    EndPreparation1 = arrayOfInputs(14)
    ScheduleThickness1 = arrayOfInputs(15)
    EndStandard1 = arrayOfInputs(16)
    PressureRating1 = arrayOfInputs(17)
    FlowDirection1 = arrayOfInputs(18)
    PortIndex2 = arrayOfInputs(19)
    Npd2 = arrayOfInputs(20)
    EndPreparation2 = arrayOfInputs(21)
    ScheduleThickness2 = arrayOfInputs(22)
    EndStandard2 = arrayOfInputs(23)
    PressureRating2 = arrayOfInputs(24)
    FlowDirection2 = arrayOfInputs(25)
    Id1 = arrayOfInputs(26)
    Id2 = arrayOfInputs(27)
    NpdUnitType = arrayOfInputs(28)
    NpdUnitType1 = arrayOfInputs(29)
    NpdUnitType2 = arrayOfInputs(30)
    
    If UBound(arrayOfInputs) >= 31 Then 'Not to fail the prior to V7 instances by accessing index 31
        parInsulationThickness = arrayOfInputs(31)
    End If
        
''---------------------------------------
        
    Dim VH          As Double
    Dim VW          As Double
    Dim HWD         As Double
    Dim HWA         As Double
    Dim POD         As Double
    Dim FTK         As Double
    Dim fd          As Double
    Dim fdmin       As Double
    Dim lines       As Collection
    Dim oLine       As IngrGeom3D.Line3d
    Dim oArc        As IngrGeom3D.Arc3d
    Dim objCStr     As IngrGeom3D.ComplexString3d
    
    Dim oStPoint     As AutoMath.DPosition
    Dim oEnPoint     As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim ldiam       As Double
    Dim objCylinder As Object

    VH = parValveHeight
    VW = parValveWidth
    HWD = parHandwheelDiameter
    
    'HWA = parHandwheelAngle / RAD
    'we are getting the angle in radians only
    HWA = parHandwheelAngle
    
    iOutput = 0
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

' Creation of Nozzle 1

''===========================
''Construction of nozzle
''===========================
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort

    Set oNozzle = CreateRetrieveDynamicNozzle(PortIndex1, Npd, NpdUnitType, _
                                            EndPreparation, ScheduleThickness, EndStandard, _
                                            PressureRating, FlowDirection, Npd1, _
                                            NpdUnitType1, EndPreparation1, ScheduleThickness1, _
                                            EndStandard1, PressureRating1, FlowDirection1, _
                                            Id1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, CptOffset, depth)
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(PortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -(faceToFace / 2 + CptOffset(PortIndex1) - depth(PortIndex1)), _
                                                                        0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

'Place Nozzle 2
    Set oNozzle = CreateRetrieveDynamicNozzle(PortIndex2, Npd, NpdUnitType, _
                                            EndPreparation, ScheduleThickness, EndStandard, _
                                            PressureRating, FlowDirection, Npd2, _
                                            NpdUnitType2, EndPreparation2, ScheduleThickness2, _
                                            EndStandard2, PressureRating2, FlowDirection2, _
                                            Id2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, CptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

    oNozzle.Length = flangeThick(PortIndex2)

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set faceToFace / 2 + CptOffset(PortIndex2) - depth(PortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    POD = BodyOD(pipeDiam(PortIndex1), pipeDiam(PortIndex2))
    dBodyFlangeThick = BodyOD(flangeThick(PortIndex1), flangeThick(PortIndex2))
    dBodyFlangeDiam = BodyOD(flangeDiam(PortIndex1), flangeDiam(PortIndex2))
    
    If CmpDblEqual(dBodyFlangeThick, 0) Then
        FTK = POD / 6
    Else
        FTK = dBodyFlangeThick
    End If
    
    If CmpDblEqual(dBodyFlangeDiam, 0) Then
        fd = POD * 1.7
    Else
        fdmin = POD * 1.7
        fd = BodyOD(dBodyFlangeDiam, fdmin)
    End If

 ' Insert your code for output 1() valve body
 Dim pnts(5) As New AutoMath.DPosition
 Dim objBspline  As IngrGeom3D.BSplineCurve3d
 Dim dtmp1 As Double
 Dim dtmp2 As Double
 dtmp1 = parValveWidth / 2 - flangeThick(PortIndex1)
 dtmp2 = parValveWidth / 2 - flangeThick(PortIndex2)
 pnts(0).Set -dtmp1, 0, pipeDiam(PortIndex1) / 2
 pnts(1).Set -dtmp1 * 0.7, 0, pnts(0).z
 pnts(2).Set -dtmp1 / 2, 0, fd * 0.6
 pnts(3).Set dtmp2 / 2, 0, fd * 0.6
 pnts(4).Set dtmp2 * 0.7, 0, pipeDiam(PortIndex2) / 2
 pnts(5).Set dtmp2, 0, pnts(4).z
 
 Set objBspline = PlaceTrBspline(3, pnts)

 Dim oCentPoint   As AutoMath.DPosition
 Set oCentPoint = New AutoMath.DPosition
 
 Dim oAxis        As AutoMath.DVector
 Set oAxis = New AutoMath.DVector
 
 Dim objRevolution  As Object
 oCentPoint.Set 0, 0, 0
 oAxis.Set 1, 0, 0
 Set objRevolution = PlaceRevolution(m_OutputColl, objBspline, _
                                    oAxis, oCentPoint, 2 * PI, _
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
    
    pnts1(1).Set 0, FTK, POD / 2
    pnts1(2).Set 0, fd / 2 + FTK, POD / 2
    pnts1(3).Set 0, fd / 2 + FTK, fd / 2
    pnts1(4).Set 0, fd / 2 + FTK * 3, fd / 2
    pnts1(5).Set 0, fd / 2 + FTK * 3, POD / 2
    pnts1(6).Set 0, fd / 2 + FTK * 3 + POD / 2, FTK / 2
    pnts1(7).Set 0, fd / 2 + FTK * 3, 0
    
    Set lines = New Collection
    For i = 1 To 4
        Set oLine = PlaceTrLine(pnts1(i), pnts1(i + 1))
        lines.Add oLine
    Next i
   
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

    oCentPoint.Set 0, 0, 0
    oAxis.Set 0, 1, 0
    Set objRevolution = PlaceRevolution(m_OutputColl, objCStr, _
                                    oAxis, oCentPoint, 2 * PI, _
                                    False)
    
    Set objCStr = Nothing

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objRevolution = Nothing
    
    
    
    
' Insert your code for output 3() create the vertical screw
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, VH - HWD, 0
    ldiam = parHandwheelDiameter / 15
    
    Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, _
                                    oEnPoint, ldiam, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    


' Insert your code for output 4() create the verticle rectangular edge
    Dim xoff As Double
    xoff = POD / 30
    
    pnts1(1).Set xoff, fd / 2 + FTK * 3, POD / 2
    pnts1(2).Set xoff, VH - HWD * 1.2, POD / 2
    pnts1(3).Set xoff, VH - HWD, POD / 2.5
    pnts1(4).Set xoff, VH - HWD, POD / 3
    pnts1(5).Set xoff, fd / 2 + FTK * 4, POD / 3
    pnts1(6).Set xoff, fd / 2 + FTK * 4, -POD / 3
    pnts1(7).Set xoff, VH - HWD, -POD / 3
    pnts1(8).Set xoff, VH - HWD, -POD / 2.5
    pnts1(9).Set xoff, VH - HWD * 1.2, -POD / 2
    pnts1(10).Set xoff, fd / 2 + FTK * 3, -POD / 2

    Set lines = New Collection
    For i = 1 To 9
        Set oLine = PlaceTrLine(pnts1(i), pnts1(i + 1))
        lines.Add oLine
    Next i
    Set oLine = PlaceTrLine(pnts1(10), pnts1(1))
    lines.Add oLine
    oStPoint.Set pnts1(1).x, pnts1(1).y, pnts1(1).z
    Set objCStr = PlaceTrCString(oStPoint, lines)
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
    xoff = POD / 5
    
    
' Create the other support member
    
    pnts1(1).Set xoff, fd / 2 + FTK * 3, POD / 2.7
    pnts1(2).Set xoff, VH - HWD, POD / 2.7
    pnts1(3).Set xoff, VH - HWD, POD / 3
    pnts1(4).Set xoff, fd / 2 + FTK * 4, POD / 3
    pnts1(5).Set xoff, fd / 2 + FTK * 4, -POD / 3
    pnts1(6).Set xoff, VH - HWD, -POD / 3
    pnts1(7).Set xoff, VH - HWD, -POD / 2.7
    pnts1(8).Set xoff, fd / 2 + FTK * 3, -POD / 2.7

    Set lines = New Collection
    For i = 1 To 7
        Set oLine = PlaceTrLine(pnts1(i), pnts1(i + 1))
        lines.Add oLine
    Next i
    
    Set oLine = PlaceTrLine(pnts1(8), pnts1(1))
    lines.Add oLine
    
    oStPoint.Set pnts1(1).x, pnts1(1).y, pnts1(1).z
    
    Set objCStr = PlaceTrCString(oStPoint, lines)
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
    For i = 1 To 7
        Set oLine = PlaceTrLine(RPS(i), RPS(i + 1))
        lines.Add oLine
    Next i
    
'''    Set oLine = PlaceTrLine(RPS(8), RPS(1))
'''    lines.Add oLine
    oStPoint.Set RPS(1).x, RPS(1).y, RPS(1).z
    Set objCStr = PlaceTrCString(oStPoint, lines)
    
    Set oLine = Nothing
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
    
    oCentPoint.Set 0, 0, 0
    oAxis.Set 0, 1, 0

    Set objRevolution = PlaceRevolution(m_OutputColl, objCStr, _
                            oAxis, oCentPoint, 2 * PI, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objRevolution = Nothing


' Spindle cone
    Dim objCone  As Object
    oStPoint.Set 0, VH - HWD / 2, 0
    oEnPoint.Set Sin(HWA) * 1.1 * fd, VH - HWD / 2, Cos(HWA) * 1.1 * fd
    Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, HWD / 5, HWD / 15, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing

' handwheel spoke 1
    Dim BP As New DPosition
    Dim TP As New DPosition
    Dim objCyl  As Object
    TP.Set Sin(HWA) * fd, VH, Cos(HWA) * fd
    oEnPoint.Set Sin(HWA) * fd, VH - HWD, Cos(HWA) * fd
    BP.Set Sin(HWA) * fd, VH - HWD / 2, Cos(HWA) * fd
    Set objCyl = PlaceCylinder(m_OutputColl, TP, oEnPoint, HWD / 25, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing

' handwheel spoke 2
    oStPoint.Set BP.x - Sin(HWA + 90 / RAD) * HWD / 2, BP.y, BP.z - Cos(HWA + 90 / RAD) * HWD / 2
    oEnPoint.Set BP.x + Sin(HWA + 90 / RAD) * HWD / 2, BP.y, BP.z + Cos(HWA + 90 / RAD) * HWD / 2
    Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, HWD / 25, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing

' handwheel torus
    Dim objTorus    As Object
    Dim CP As New DPosition
    CP.Set BP.x, BP.y, BP.z
    oAxis.Set Sin(HWA), 0, Cos(HWA)
    Set objTorus = PlaceTorus(m_OutputColl, CP, oAxis, HWD / 2, HWD / 20)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTorus
    Set objTorus = Nothing

'handwheel knob
    oEnPoint.Set TP.x + Sin(HWA) * HWD / 5, TP.y, TP.z + Cos(HWA) * HWD / 5
    Set objCyl = PlaceCylinder(m_OutputColl, TP, oEnPoint, HWD / 25, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing


    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing


    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

'Public Function GetCatalogDBConnection() As IJDConnection
'    Const METHOD = "GetCatalogDBConnection"
'    On Error GoTo ErrorHandler
'
'    Dim oTrader As IMSTrader.Trader
'    Set oTrader = New Trader
'    Dim oWorkingSet As WorkingSetLibrary.IJDWorkingSet
'    Set oWorkingSet = oTrader.Service("WorkingSet", "")
'    Dim strCatlogDB As String
'    Dim oAppCtx As IJApplicationContext
'    Set oAppCtx = oTrader.Service("ApplicationContext", "")
'    strCatlogDB = oAppCtx.DBTypeConfiguration.get_DataBaseFromDBType("Catalog")
'    Set GetCatalogDBConnection = oWorkingSet.Item(strCatlogDB)
'    Set oTrader = Nothing
'    Set oWorkingSet = Nothing
'    Set oAppCtx = Nothing
'
'    Exit Function
'ErrorHandler:
'    ReportUnanticipatedError MODULE, METHOD
'End Function

Public Function GetCatalogDBResourceManager() As IUnknown
    Const METHOD = "GetCatalogDBResourceManager"
    On Error GoTo ErrorHandler
    Dim oMidctx As IJMiddleContext
    Set oMidctx = New GSCADMiddleContextProj.GSCADMiddleContext
    Set GetCatalogDBResourceManager = oMidctx.GetResourceManager("Catalog")
    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
End Function



