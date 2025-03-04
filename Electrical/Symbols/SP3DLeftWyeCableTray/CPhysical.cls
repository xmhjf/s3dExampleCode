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
'   Author:          BG
'   Creation Date:  Thursday, Mar 7 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' SP3DLeftWyeCableTray is placed with port-1 on negative X, Port2 on positive X
' and Port-3 towards positive Y axis respectively. The Geometry is formed using four outputs:
' 1. Bottom of Wye
' 2. Port1 to Port3 Branch
' 3. Port2 and Port3 Junction
' 4. Right side Plane
' Outputs 5,6 & 7 are Cable Tray Ports created by using 'CreateCableTrayPort' function.
'
'   Change History:
'   dd.mmm.yyyy          who                    change description
'   -----------         -----                    ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.May.2006           kkc               Code added to include LeftWye Cable Tray of TJ Cope
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI    As Double


Private Sub Class_Initialize()
      PI = 4 * Atn(1)
'''
      

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
        
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ActualWidth As Double
    Dim ActualDepth As Double
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim oPort3 As New AutoMath.DPosition 'Port 3 center point
    Const EI_IN = 0.2032 'Length of the header upto Branch for CJ Cope
    
    Dim HD              As Double
    Dim HW              As Double
    Dim HW3              As Double
    Dim parFacetoFace As Double
    Dim ActualWidth3 As Double
    Dim Port1S(0 To 11) As Double
    Dim Port2S(0 To 11) As Double
    Dim Port3S(0 To 11) As Double
    Dim parAngle As Double
    Dim lCableTrayPartDataBasis As Long
    
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim iOutput     As Double
    Dim ObjCurvedBranch As Object
    Dim ObjVShapeBranch As Object
    Dim parTangentLength As Double
    Dim parFacetoTangent As Double
    Dim parBendRadius As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoTangent = arrayOfInputs(2)
       
    iOutput = 0
    parAngle = PI / 4       'Angle  45 degree
 ' Insert your code for output 1(Wye Bottom Portion)
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    parFacetoFace = oTrayPart.Length
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    
    
    'parTangentLength = 0.0762
    parTangentLength = oTrayPart.TangentLength
    parBendRadius = oTrayPart.BendRadius
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, ActualWidth, ActualDepth)
    Call RetrieveCableTrayPortProperties(3, oPartFclt, ActualWidth3, ActualDepth)
    HD = ActualDepth / 2
    HW = ActualWidth / 2
    HW3 = ActualWidth3 / 2
    'Calculating the Angle Between the starting point of curve and the co-ordinate origin
    'Angle Between Starting Point of Curve Near Port-1 and Origin (AngleP1andO) Let us take "A"
    'Ratio of (R+HW3)/(R+HW) - (RatioWidth) CableTray Center from Center Point of Curve
    'Cos(45-A)/Cos(A)=(R+HW3)/(R+HW)
    'Cos45*Cos(A)+Sin(45)*SinA=((R+HW3)/(R+HW))*Cos(A)
    Dim RatioWidth As Double
    Dim SqrTwo As Double
    Dim AngleP1andO As Double
    Dim AngleP2andO As Double
    RatioWidth = (parBendRadius + HW3) / (parBendRadius + HW)
    SqrTwo = Sqr(2)
    AngleP1andO = Atn(RatioWidth * SqrTwo - 1)
    AngleP2andO = parAngle - AngleP1andO
    
    'Port-1 Position
    If lCableTrayPartDataBasis <= 1 Then
     oPort1.Set -(parBendRadius + HW) * Tan(AngleP1andO) - parTangentLength, 0, 0
    ElseIf lCableTrayPartDataBasis = 5 Then
     oPort1.Set -parFacetoTangent, 0, 0
    End If
    'Port-1 Top Point on Positive Y Direction
    Port1S(0) = oPort1.x
    Port1S(1) = oPort1.y + HW
    Port1S(2) = oPort1.z + HD
    'Port-1 Bottom Point on Positive Y Direction
    Port1S(3) = oPort1.x
    Port1S(4) = oPort1.y + HW
    Port1S(5) = oPort1.z - HD
    'Port-1 Bottom Point on Negative Y Direction
    Port1S(6) = oPort1.x
    Port1S(7) = oPort1.y - HW
    Port1S(8) = oPort1.z - HD
    'Port-1 Top Point on Negative Y Direction
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD
    If lCableTrayPartDataBasis <= 1 Then
    'Port-2 Position
     oPort2.Set parFacetoFace - (parBendRadius + HW) * Tan(AngleP1andO) - parTangentLength, 0, 0
    ElseIf lCableTrayPartDataBasis = 5 Then
     oPort2.Set parFacetoFace - parFacetoTangent, 0, 0
    End If
    'Port-1 Top Point on Positive Y Direction
    Port2S(0) = oPort2.x
    Port2S(1) = oPort2.y + HW
    Port2S(2) = oPort2.z + HD
    'Port-1 Bottom Point on Positive Y Direction
    Port2S(3) = oPort2.x
    Port2S(4) = oPort2.y + HW
    Port2S(5) = oPort2.z - HD
    'Port-1 Bottom Point on Negative Y Direction
    Port2S(6) = oPort2.x
    Port2S(7) = oPort2.y - HW
    Port2S(8) = oPort2.z - HD
    'Port-1 Top Point on Negative Y Direction
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD

    'Port-3 Position
    'Calculating Distance along the branch From the Origin upto Port3 (DFOP3)
    Dim DFOP3  As Double
    Dim BCSP(0 To 2) As Double
    Dim JPBP(0 To 2) As Double
    If lCableTrayPartDataBasis <= 1 Then
     DFOP3 = (parBendRadius + HW3) * Tan(AngleP2andO) + parFacetoTangent
     oPort3.Set DFOP3 * Cos(parAngle), DFOP3 * Sin(parAngle), 0
    ElseIf lCableTrayPartDataBasis = 5 Then
     oPort3.Set parFacetoFace - parFacetoTangent, parFacetoFace - parFacetoTangent, 0
    End If
    'Port-3 Top Point on Positive Y Direction
    Port3S(0) = oPort3.x - HW3 * Sin(parAngle)
    Port3S(1) = oPort3.y + HW3 * Cos(parAngle)
    Port3S(2) = oPort3.z + HD
    'Port-3 Bottom Point on Positive Y Direction
    Port3S(3) = oPort3.x - HW3 * Sin(parAngle)
    Port3S(4) = oPort3.y + HW3 * Cos(parAngle)
    Port3S(5) = oPort3.z - HD
    'Port-3 Bottom Point on Negative Y Direction
    Port3S(6) = oPort3.x + HW3 * Sin(parAngle)
    Port3S(7) = oPort3.y - HW3 * Cos(parAngle)
    Port3S(8) = oPort3.z - HD
    'Port-3 Top Point on Negative Y Direction
    Port3S(9) = oPort3.x + HW3 * Sin(parAngle)
    Port3S(10) = oPort3.y - HW3 * Cos(parAngle)
    Port3S(11) = oPort3.z + HD
    'Branch Curve Start Point near Port-1 BCSP
    If lCableTrayPartDataBasis <= 1 Then
    Dim BCEP(0 To 2) As Double
    Dim BCCP(0 To 2) As Double
    BCSP(0) = Port1S(3) + parTangentLength
    BCSP(1) = Port1S(4)
    BCSP(2) = Port1S(5)
    'Branch Curve End Point near Port-3 BCEP
    BCEP(0) = Port3S(3) - parFacetoTangent * Cos(parAngle)
    BCEP(1) = Port3S(4) - parFacetoTangent * Sin(parAngle)
    BCEP(2) = Port3S(5)
    'Branch Curve Center Point BCCP
    BCCP(0) = BCSP(0)
    BCCP(1) = BCSP(1) + parBendRadius
    BCCP(2) = BCSP(2)
    
    ElseIf lCableTrayPartDataBasis = 5 Then
    BCSP(0) = Port1S(3) + EI_IN
    BCSP(1) = Port1S(4)
    BCSP(2) = Port1S(5)
    End If
    
    
    'Junction Point Between Port-2 and Port-3 (JPBP)
    JPBP(0) = HW + HW3 / Sin(parAngle)
    JPBP(1) = HW
    JPBP(2) = -HD
    '   Construct Top Tangent line Near Port-1 on +ive Y-axis
    Dim oTTLineP1           As IngrGeom3D.Line3d
    
    Set oTTLineP1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), _
                                Port1S(5), BCSP(0), BCSP(1), BCSP(2))
    
    'Construct the Branch Curve
    If lCableTrayPartDataBasis <= 1 Then
    Dim objBRA  As IngrGeom3D.Arc3d
    Dim SBRA   As New AutoMath.DPosition
    Dim EBRA   As New AutoMath.DPosition
    Dim CBRA   As New AutoMath.DPosition
    SBRA.Set BCSP(0), BCSP(1), BCSP(2)
    EBRA.Set BCEP(0), BCEP(1), BCEP(2)
    CBRA.Set BCCP(0), BCCP(1), BCCP(2)
    Set objBRA = PlaceTrArcByCenter(SBRA, EBRA, CBRA)
    '   Construct Top Tangent line Near Port-3 on +ive Y-axis
    Dim oTTLineP3           As IngrGeom3D.Line3d
    Set oTTLineP3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BCEP(0), BCEP(1), BCEP(2), _
                                                                                                    Port3S(3), Port3S(4), Port3S(5))
    
    ElseIf lCableTrayPartDataBasis = 5 Then
    Dim oTTTjCope           As IngrGeom3D.Line3d
    Set oTTTjCope = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BCSP(0), BCSP(1), BCSP(2), _
                               Port3S(3), Port3S(4), Port3S(5))
    End If
    
    
    '   Construct line along Port-3
    Dim oLineAP3           As IngrGeom3D.Line3d
    Set oLineAP3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), _
                                                                                                    Port3S(6), Port3S(7), Port3S(8))
    '   Construct line Between Port-3 Junction Point
    Dim oLineP3JP           As IngrGeom3D.Line3d
    Set oLineP3JP = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), _
                                                                                                    JPBP(0), JPBP(1), JPBP(2))
    '   Construct line Between Junction Point and Port-2 +ive Side
    Dim oLineJPP2           As IngrGeom3D.Line3d
    Set oLineJPP2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, JPBP(0), JPBP(1), JPBP(2), _
                                                                                                    Port2S(3), Port2S(4), Port2S(5))
    '   Construct line Along Port-2
    Dim oLineAP2           As IngrGeom3D.Line3d
    Set oLineAP2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), _
                                                                                                Port2S(6), Port2S(7), Port2S(8))
    '   Construct Header Line Between Port-1 and Port-2
    Dim oLineHeader           As IngrGeom3D.Line3d
    Set oLineHeader = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), _
                                                                                                Port1S(6), Port1S(7), Port1S(8))
        '   Construct Line Along Port-1
    Dim oLineAP1           As IngrGeom3D.Line3d
    Set oLineAP1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), _
                                                                                                Port1S(3), Port1S(4), Port1S(5))
                                                                                                
    '   Construct entire Bottom header and branch curves
    Dim obottomcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Dim objHcurves As IngrGeom3D.ComplexString3d
    Set obottomcurves = New Collection
    
    

    obottomcurves.Add oTTLineP1

    If lCableTrayPartDataBasis <= 1 Then
        obottomcurves.Add objBRA
        obottomcurves.Add oTTLineP3
    ElseIf lCableTrayPartDataBasis = 5 Then
        obottomcurves.Add oTTTjCope
    End If
    obottomcurves.Add oLineAP3
    obottomcurves.Add oLineP3JP
    obottomcurves.Add oLineJPP2
    obottomcurves.Add oLineAP2
    obottomcurves.Add oLineHeader
    obottomcurves.Add oLineAP1
    
    
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set objHBcurves = PlaceTrCString(StartBC, obottomcurves)
    Dim TrayBottom As IngrGeom3D.Plane3d
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    Set TrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                                Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call TrayBottom.AddBoundary(objHBcurves)
    'Remove cable tray bottom Header and Branch lines
    Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
                                                                                                    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), TrayBottom
    Set TrayBottom = Nothing
    
  'Insert your code for output 2(Branch Curved Portion)
    Dim objBrCurve         As IngrGeom3D.ComplexString3d
    Dim oBrCurve           As Collection
    Set oBrCurve = New Collection
    oBrCurve.Add oTTLineP1
    If lCableTrayPartDataBasis <= 1 Then
     oBrCurve.Add objBRA
     oBrCurve.Add oTTLineP3
    ElseIf lCableTrayPartDataBasis = 5 Then
     oBrCurve.Add oTTTjCope
    End If
    Set objBrCurve = PlaceTrCString(StartBC, oBrCurve)
    Set ObjCurvedBranch = PlaceProjection(m_OutputColl, objBrCurve, oDirProj, ActualDepth, True)
    Set objBrCurve = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCurvedBranch
    Set ObjCurvedBranch = Nothing
    
 ' Insert your code for output 3(V Shape Branch between2&3)
    StartBC.Set Port3S(6), Port3S(7), Port3S(8)
    Dim objVShapeCurves         As IngrGeom3D.ComplexString3d
    Dim oVShapeCurves           As Collection
    Set oVShapeCurves = New Collection
    oVShapeCurves.Add oLineP3JP
    oVShapeCurves.Add oLineJPP2
    Set objVShapeCurves = PlaceTrCString(StartBC, oVShapeCurves)
    Set ObjVShapeBranch = PlaceProjection(m_OutputColl, objVShapeCurves, oDirProj, ActualDepth, True)
    Set objVShapeCurves = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVShapeBranch
    Set ObjVShapeBranch = Nothing
    Set StartBC = Nothing
    Set oDirProj = Nothing
' Insert your code for output 4(Header between port-1 and port-2)
    Dim HSide(0 To 11) As Double
    Dim HeaderSide As IngrGeom3D.Plane3d
    HSide(0) = Port2S(9)
    HSide(1) = Port2S(10)
    HSide(2) = Port2S(11)
    
    HSide(3) = Port2S(6)
    HSide(4) = Port2S(7)
    HSide(5) = Port2S(8)
    
    HSide(6) = Port1S(6)
    HSide(7) = Port1S(7)
    HSide(8) = Port1S(8)
    
    HSide(9) = Port1S(9)
    HSide(10) = Port1S(10)
    HSide(11) = Port1S(11)
    
    Set HeaderSide = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, HSide)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), HeaderSide
    
    Set HeaderSide = Nothing
    Set oTTLineP1 = Nothing
    If lCableTrayPartDataBasis <= 1 Then
    Set objBRA = Nothing
    Set oTTLineP3 = Nothing
    ElseIf lCableTrayPartDataBasis = 5 Then
    Set oTTTjCope = Nothing
    End If
    Set oLineAP3 = Nothing
    Set oLineP3JP = Nothing
    Set oLineJPP2 = Nothing
    Set oLineAP2 = Nothing
    Set oLineHeader = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To obottomcurves.Count
        obottomcurves.Remove 1
    Next iCount
    Set obottomcurves = Nothing
    
    For iCount = 1 To oBrCurve.Count
        oBrCurve.Remove 1
    Next iCount
    Set oBrCurve = Nothing
    
    For iCount = 1 To oVShapeCurves.Count
        oVShapeCurves.Remove 1
    Next iCount
    Set oVShapeCurves = Nothing
    
' Place Port 1
    
    'Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    
    'Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    'oPlacePoint.Set -faceToFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    '    Set oPlacePoint = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Port 2
    '    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition

'    oPlacePoint.Set 0, (parBendRadius + parActualWidth / 2), 0
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1
    
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Port 3
    '    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition

'    oPlacePoint.Set 0, (parBendRadius + parActualWidth / 2), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    oRadialOrient.Set 0, 0, 1
    
    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort3 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oGeomFactory = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

