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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Mar 7 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbolfor this aspect
' SP3DRightWyeCableTray is placed with port-1 on negative X, Port2 on positive X
' and Port-3 towards negative Y axis respectively. The Geometry is formed using four outputs:
' 1. Bottom of Wye
' 2. Port1 to Port3 Branch
' 3. Port2 and Port3 Junction
' 4. Left side Plane
' Outputs 5,6 & 7 are Cable Tray Ports created by using 'CreateCableTrayPort' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.May.2006           kkc               CR-83908: Code added to include RightWye Cable Tray of TJ Cope
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim ObjTrayBottom      As Object
    Dim ObjP1P3Branch      As Object
    Dim ObjP2P3Junction    As Object
    Dim ObjLeftPlane       As Object
    
    Dim parFacetoTangent  As Double
    
    Dim iOutput            As Double
    Dim ActualWidth     As Double
    Dim ActualDepth     As Double
    Dim ActualWidth3    As Double
    Dim parLength          As Double
    Dim parTangentLength   As Double
    Dim parBendRadius      As Double
    Dim HD                 As Double
    Dim HW                 As Double
    Dim HW3                As Double
    Dim LPlane(0 To 11)    As Double
    Dim Port1S(0 To 11)    As Double
    Dim Port2S(0 To 11)    As Double
    Dim Port3S(0 To 11)    As Double
    Dim oPort1             As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2             As New AutoMath.DPosition 'Port 2 center point
    Dim oPort3             As New AutoMath.DPosition 'Port 3 center point
    Const EI_IN = 0.2032 'Length of the header upto Branch for CJ Cope
    Dim lCableTrayPartDataBasis As Long

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoTangent = arrayOfInputs(2)
'   Retrieve Part properties: Length, BendRadius, TangentLength
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    parLength = oTrayPart.Length
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    parBendRadius = oTrayPart.BendRadius
'   parTangentLength = 0.0762
    parTangentLength = oTrayPart.TangentLength
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
   
    
    iOutput = 0

    Call RetrieveCableTrayPortProperties(1, oPartFclt, ActualWidth, ActualDepth)
    Call RetrieveCableTrayPortProperties(3, oPartFclt, ActualWidth3, ActualDepth)
    HD = ActualDepth / 2
    HW = ActualWidth / 2
    HW3 = ActualWidth3 / 2
    
'   Compute angle 'Theta' between a)Line joining bend-center to origin(between Port1 and Port3)
'   and b)the positive Y axis
    Dim parTheta As Double
    parTheta = Atn((Sqr(2) * (parBendRadius + HW3) / (parBendRadius + HW)) - 1)

' Insert your code for output 1(TrayBottom)
'   Prepare port1 points
    If lCableTrayPartDataBasis <= 1 Then
     oPort1.Set -(parBendRadius + HW) * Tan(parTheta) - parTangentLength, 0, 0
    ElseIf lCableTrayPartDataBasis = 5 Then
     oPort1.Set -parFacetoTangent, 0, 0
    End If

'   Tray top edge on Positive Y-Axis Side
    Port1S(0) = oPort1.x
    Port1S(1) = oPort1.y + HW
    Port1S(2) = oPort1.z + HD

'   Tray bottom on Positive Y-Axis Side
    Port1S(3) = oPort1.x
    Port1S(4) = oPort1.y + HW
    Port1S(5) = oPort1.z - HD

'   Tray bottom on Negative Y-Axis Side
    Port1S(6) = oPort1.x
    Port1S(7) = oPort1.y - HW
    Port1S(8) = oPort1.z - HD
    
'   Tray top on Negative Y-Axis Side
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD

'   Prepare port2 points
    If lCableTrayPartDataBasis <= 1 Then
     oPort2.Set -(parBendRadius + HW) * Tan(parTheta) - parTangentLength + parLength, 0, 0
    ElseIf lCableTrayPartDataBasis = 5 Then
     oPort2.Set parLength - parFacetoTangent, 0, 0
    End If
'   Tray top edge on Positive Y-Axis Side
    Port2S(0) = oPort2.x
    Port2S(1) = oPort2.y + HW
    Port2S(2) = oPort2.z + HD

'   Tray bottom on Positive Y-Axis Side
    Port2S(3) = oPort2.x
    Port2S(4) = oPort2.y + HW
    Port2S(5) = oPort2.z - HD

'   Tray bottom on Negative Y-Axis Side
    Port2S(6) = oPort2.x
    Port2S(7) = oPort2.y - HW
    Port2S(8) = oPort2.z - HD
    
'   Tray top on Negative Y-Axis Side
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD

'   Prepare port3 points
    If lCableTrayPartDataBasis <= 1 Then
        Dim parFace3toCenter    As Double
        parFace3toCenter = (parBendRadius + HW3) * Tan(PI / 4 - parTheta) + parFacetoTangent
        oPort3.Set parFace3toCenter * Cos(PI / 4), -parFace3toCenter * Sin(PI / 4), 0
    ElseIf lCableTrayPartDataBasis = 5 Then
        oPort3.Set parLength - parFacetoTangent, parFacetoTangent - parLength, 0
    End If

'   Tray top edge towards Positive Y-Axis Side
    Port3S(0) = oPort3.x + HW3 * Sin(PI / 4)
    Port3S(1) = oPort3.y + HW3 * Cos(PI / 4)
    Port3S(2) = oPort3.z + HD

'   Tray bottom towards Positive Y-Axis Side
    Port3S(3) = oPort3.x + HW3 * Sin(PI / 4)
    Port3S(4) = oPort3.y + HW3 * Cos(PI / 4)
    Port3S(5) = oPort3.z - HD

'   Tray bottom towards Negative Y-Axis Side
    Port3S(6) = oPort3.x - HW3 * Sin(PI / 4)
    Port3S(7) = oPort3.y - HW3 * Cos(PI / 4)
    Port3S(8) = oPort3.z - HD
    
'   Tray top towards Negative Y-Axis Side
    Port3S(9) = oPort3.x - HW3 * Sin(PI / 4)
    Port3S(10) = oPort3.y - HW3 * Cos(PI / 4)
    Port3S(11) = oPort3.z + HD
    
'   Prepare branch curve bottom points between Port1 and Port3
    If lCableTrayPartDataBasis <= 1 Then
        Dim BCbottomStartPt(0 To 2)     As Double
        Dim BCbottomEndPt(0 To 2)       As Double
        Dim BCbottomCenterPt(0 To 2)    As Double
    
        BCbottomStartPt(0) = oPort1.x + parTangentLength
        BCbottomStartPt(1) = oPort1.y - HW
        BCbottomStartPt(2) = oPort1.z - HD
        
        BCbottomEndPt(0) = oPort3.x - HW3 * Sin(PI / 4) - parFacetoTangent * Cos(PI / 4)
        BCbottomEndPt(1) = oPort3.y - HW3 * Cos(PI / 4) + parFacetoTangent * Sin(PI / 4)
        BCbottomEndPt(2) = oPort3.z - HD
        
        BCbottomCenterPt(0) = oPort1.x + parTangentLength
        BCbottomCenterPt(1) = -HW - parBendRadius
        BCbottomCenterPt(2) = oPort1.z - HD
    
    ElseIf lCableTrayPartDataBasis = 5 Then
        BCbottomStartPt(0) = oPort1.x + EI_IN
        BCbottomStartPt(1) = oPort1.y - HW
        BCbottomStartPt(2) = oPort1.z - HD
    End If
    
'   Prepare tray bottom junction point between Port2 and Port3
    Dim P2P3JunctionPt(0 To 2) As Double
    P2P3JunctionPt(0) = HW + HW3 / Sin(PI / 4)
    P2P3JunctionPt(1) = -HW
    P2P3JunctionPt(2) = oPort3.z - HD
    
'   Construct Tray bottom:
'   Construct Port1 Bottom line
    Dim oP1BottomLine As IngrGeom3D.Line3d
    Set oP1BottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), _
                              Port1S(6), Port1S(7), Port1S(8))
    Dim oP1RHSTangentLine As IngrGeom3D.Line3d
    Set oP1RHSTangentLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), _
                                    BCbottomStartPt(0), BCbottomStartPt(1), BCbottomStartPt(2))
    
    If lCableTrayPartDataBasis = 5 Then
    '   Construct Port1 Right hand side Tangent line
        
        Dim oTTTjCope           As IngrGeom3D.Line3d
        Set oTTTjCope = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BCbottomStartPt(0), BCbottomStartPt(1), BCbottomStartPt(2), _
                                Port3S(6), Port3S(7), Port3S(8))
    ElseIf lCableTrayPartDataBasis <= 1 Then
    '   Construct branch curve
        Dim objBArc  As IngrGeom3D.Arc3d
        Dim SBArc   As New AutoMath.DPosition
        Dim EBArc   As New AutoMath.DPosition
        Dim CBArc   As New AutoMath.DPosition
        SBArc.Set BCbottomStartPt(0), BCbottomStartPt(1), BCbottomStartPt(2)
        EBArc.Set BCbottomEndPt(0), BCbottomEndPt(1), BCbottomEndPt(2)
        CBArc.Set BCbottomCenterPt(0), BCbottomCenterPt(1), BCbottomCenterPt(2)
        Set objBArc = PlaceTrArcByCenter(SBArc, EBArc, CBArc)
        
    '   Construct Port3 Right hand side line
        Dim oP3RHSLine As IngrGeom3D.Line3d
        Set oP3RHSLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BCbottomEndPt(0), BCbottomEndPt(1), BCbottomEndPt(2), _
                                        Port3S(6), Port3S(7), Port3S(8))
    End If
'   Construct Port3 Bottom line
    Dim oP3BottomLine As IngrGeom3D.Line3d
    Set oP3BottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), _
                                    Port3S(3), Port3S(4), Port3S(5))
                                    
'   Construct Port3 Left hand side line which joins to the junction point
    Dim oP3LHSLine As IngrGeom3D.Line3d
    Set oP3LHSLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), _
                                    P2P3JunctionPt(0), P2P3JunctionPt(1), P2P3JunctionPt(2))

'   Construct Port2 Right hand side line which joins to the junction point
    Dim oP2RHSLine As IngrGeom3D.Line3d
    Set oP2RHSLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, P2P3JunctionPt(0), P2P3JunctionPt(1), P2P3JunctionPt(2), _
                                    Port2S(6), Port2S(7), Port2S(8))
                
'   Construct Port2 Bottom line
    Dim oP2BottomLine As IngrGeom3D.Line3d
    Set oP2BottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), _
                                    Port2S(3), Port2S(4), Port2S(5))
                            
'   Construct Port1 to Port2 Left hand side bottom line
    Dim oP1P2LHSbottomLine As IngrGeom3D.Line3d
    Set oP1P2LHSbottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), _
                                    Port1S(3), Port1S(4), Port1S(5))

'   Construct complete bottom boundary curves
    Dim obottomcurves       As Collection
    Dim ObjCBottomCurves    As IngrGeom3D.ComplexString3d
    Set obottomcurves = New Collection
    obottomcurves.Add oP1BottomLine
    obottomcurves.Add oP1RHSTangentLine
    If lCableTrayPartDataBasis <= 1 Then
        obottomcurves.Add objBArc
        obottomcurves.Add oP3RHSLine
    ElseIf lCableTrayPartDataBasis = 5 Then
        obottomcurves.Add oTTTjCope
    End If
    
    obottomcurves.Add oP3BottomLine
    obottomcurves.Add oP3LHSLine
    obottomcurves.Add oP2RHSLine
    obottomcurves.Add oP2BottomLine
    obottomcurves.Add oP1P2LHSbottomLine
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set ObjCBottomCurves = PlaceTrCString(StartBC, obottomcurves)
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
'   Create the Wye bottom
    Set ObjTrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjTrayBottom.AddBoundary(ObjCBottomCurves)
'   Remove cable tray bottom
    Dim ObjBcurves As IJDObject
    Set ObjBcurves = ObjCBottomCurves
    ObjBcurves.Remove
    Set ObjBcurves = Nothing
    Set StartBC = Nothing
    

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrayBottom
    Set ObjTrayBottom = Nothing
    
' Insert your code for output 2(P1P3Branch)
    Dim ObjP1P3RHSbottomcurve As IngrGeom3D.ComplexString3d
    Dim oP1P3RHSbottomcurve   As Collection
    Set oP1P3RHSbottomcurve = New Collection
    If lCableTrayPartDataBasis <= 1 Then
    oP1P3RHSbottomcurve.Add oP1RHSTangentLine
    oP1P3RHSbottomcurve.Add objBArc
    oP1P3RHSbottomcurve.Add oP3RHSLine
    
    ElseIf lCableTrayPartDataBasis = 5 Then
    oP1P3RHSbottomcurve.Add oP1RHSTangentLine
    oP1P3RHSbottomcurve.Add oTTTjCope
    End If
    
    Dim StartP1P3BC   As New AutoMath.DPosition
    StartP1P3BC.Set Port1S(6), Port1S(7), Port1S(8)
    Set ObjP1P3RHSbottomcurve = PlaceTrCString(StartP1P3BC, oP1P3RHSbottomcurve)
    Set ObjP1P3Branch = PlaceProjection(m_OutputColl, ObjP1P3RHSbottomcurve, oDirProj, ActualDepth, True)
    Set ObjP1P3RHSbottomcurve = Nothing
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjP1P3Branch
    Set ObjP1P3Branch = Nothing
    Set StartP1P3BC = Nothing

' Insert your code for output 3(P2P3Junction)
    Dim ObjP2P3JunctionCurve As IngrGeom3D.ComplexString3d
    Dim oP2P3JunctionCurve   As Collection
    Set oP2P3JunctionCurve = New Collection
    oP2P3JunctionCurve.Add oP3LHSLine
    oP2P3JunctionCurve.Add oP2RHSLine
    Dim StartP2P3BC   As New AutoMath.DPosition
    StartP2P3BC.Set Port3S(3), Port3S(4), Port3S(5)
    Set ObjP2P3JunctionCurve = PlaceTrCString(StartP2P3BC, oP2P3JunctionCurve)
    Set ObjP2P3Junction = PlaceProjection(m_OutputColl, ObjP2P3JunctionCurve, oDirProj, ActualDepth, True)
    Set ObjP2P3JunctionCurve = Nothing
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjP2P3Junction
    Set ObjP2P3Junction = Nothing
    Set StartP2P3BC = Nothing

' Insert your code for output 4(LeftPlane)
'   Co-ordinates for Left side plane
    LPlane(0) = oPort1.x
    LPlane(1) = oPort1.y + HW
    LPlane(2) = oPort1.z - HD
    
    LPlane(3) = oPort2.x
    LPlane(4) = oPort2.y + HW
    LPlane(5) = oPort2.z - HD
    
    LPlane(6) = oPort2.x
    LPlane(7) = oPort2.y + HW
    LPlane(8) = oPort2.z + HD
    
    LPlane(9) = oPort1.x
    LPlane(10) = oPort1.y + HW
    LPlane(11) = oPort1.z + HD
    Set ObjLeftPlane = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, LPlane)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftPlane
    Set ObjLeftPlane = Nothing

' Place Nozzle 1
'   Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim ObjCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

'   oPlacePoint.Set -parLength / 2, 0, 0
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set ObjCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Nozzle 2
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set parLength / 2, 0, 0
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set ObjCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Nozzle 3
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set parLength / 2, 0, 0
    oDir.Set Cos(PI / 4), -Sin(PI / 4), 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set ObjCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort3 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing

'   Remove curves
    Set oP1BottomLine = Nothing
    Set oP1RHSTangentLine = Nothing
    
    Set oP3BottomLine = Nothing
    Set oP3LHSLine = Nothing
    Set oP2RHSLine = Nothing
    Set oP2BottomLine = Nothing
    Set oP1P2LHSbottomLine = Nothing
    If lCableTrayPartDataBasis <= 1 Then
        Set objBArc = Nothing
        Set oP3RHSLine = Nothing
    ElseIf lCableTrayPartDataBasis = 5 Then
        Set oTTTjCope = Nothing
    End If
    Dim iCount As Integer
    For iCount = 1 To obottomcurves.Count
        obottomcurves.Remove 1
    Next iCount
    Set obottomcurves = Nothing
    
    For iCount = 1 To oP1P3RHSbottomcurve.Count
        oP1P3RHSbottomcurve.Remove 1
    Next iCount
    Set oP1P3RHSbottomcurve = Nothing
    
    For iCount = 1 To oP2P3JunctionCurve.Count
        oP2P3JunctionCurve.Remove 1
    Next iCount
    Set oP2P3JunctionCurve = Nothing
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim sptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - sptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + sptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
