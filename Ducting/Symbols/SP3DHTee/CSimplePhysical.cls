VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   CSimplePhysical.cls

'   Author:         RRK
'   Creation Date:  Monday, May 14 2007
'   Description:
'   This symbol is created for Straight And Conical Tees of McGill Air Flow Corporation

'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Const INCH = 0.0254
Private PI       As Double
Private Sub Class_Initialize()
 PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CP As New AutoMath.DPosition
    Dim CV As New AutoMath.DPosition
    Dim NozzleLength As Double
    
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    Dim parPlaneOfBranch As Double
    
    Dim PortDirection As AutoMath.DVector
    Set PortDirection = New AutoMath.DVector
    Dim RadialDirection As AutoMath.DVector
    Set RadialDirection = New AutoMath.DVector

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parBWidth = arrayOfInputs(5)
    parBDepth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    parPlaneOfBranch = arrayOfInputs(8)

    iOutput = 0
    
    Dim CornerRadius As Double
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If

'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    Dim dHlength As Double
    Dim dCentFrmPort2 As Double
    Dim dBLength As Double
        
' Insert your code for output 1(Branch Body)

    Dim dTemp As Double
    
    If (lPartdatabasis <= 1) Or (lPartdatabasis = 5) Then
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            parBWidth = parWidth
        End If
        
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
        
        dHlength = parBWidth + (2 * INCH)
        
        If parPlaneOfBranch = PI / 2 Then
            dBLength = parDepth / 2 + INCH
        Else
            dBLength = parWidth / 2 + INCH
        End If
        
    ElseIf (lPartdatabasis = 10) Then
        If CmpDblGreaterthan(parBWidth, parWidth - (2 * INCH)) Then
            parBWidth = parWidth - (2 * INCH)
        End If
        
        If CmpDblGreaterthan(parBDepth, parDepth - (2 * INCH)) Then
            parBDepth = parDepth - (2 * INCH)
        End If
        
        dHlength = parBWidth + (4 * INCH)
        
        If parPlaneOfBranch = PI / 2 Then
            dBLength = parDepth / 2 + (4 * INCH)
        Else
            dBLength = parWidth / 2 + (4 * INCH)
        End If
    End If
    
    Dim dBWidth As Double
    Dim dBDepth As Double
    
    Dim dBranchStartY As Double
    Dim dBranchDia As Double
    
    If (lPartdatabasis <= 1) Or (lPartdatabasis = 5) Then
        dBranchDia = parBWidth
        dBWidth = parBWidth
        dBDepth = parBDepth
    ElseIf (lPartdatabasis = 10) Then
        dBranchDia = parBDepth + (1 * INCH)
        dBWidth = parBWidth + (1 * INCH)
        dBDepth = parBDepth + (1 * INCH)
    End If
    
    If CmpDblGreaterthan(dBranchDia, parWidth) Then
        dBranchDia = parWidth
    End If
    
    dBranchStartY = Sqr((parWidth / 2) ^ 2 - (dBranchDia / 2) ^ 2)


    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oCurve1 As Object
    Dim oCurve2 As Object
    
    Dim oHeaderStripCurve1 As Object
    Dim oHeaderStripCurve2 As Object
    Dim oHeaderTakeoffCurve1 As Object
    Dim oHeaderTakeoffCurve2 As Object
    Dim oBranchStripCurve As Object
    Dim oBranchTakeoffCurve As Object
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim objProfile As Object
    Dim PlaneofTurn As Long

    
    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector
    'create the profile for the sweep
    If parHVACShape = FlatOval Then
        If CmpDblGreaterthan(dBDepth, parDepth) Then
            dBDepth = parDepth
        End If
        
        dBranchStartY = ((parWidth - parDepth) / 2) + ((parDepth / 2) ^ 2 - (dBDepth / 2) ^ 2)
        CP.Set 0, dBranchStartY, 0
        Set oCurve1 = CreFlatOvalBranch(CP, dBWidth, dBDepth, parPlaneOfBranch)
        
        CP.Set 0, dBLength, 0
        Set oCurve2 = CreFlatOvalBranch(CP, parBWidth, parBDepth, parPlaneOfBranch)
        
        Set oBranchTakeoffCurve = CreFlatOvalBranch(CP, 1.01 * parBWidth, 1.01 * parBDepth, parPlaneOfBranch)
        
        CP.Set 0, dBLength - 0.004, 0
        Set oBranchStripCurve = CreFlatOvalBranch(CP, 1.06 * parBWidth, 1.06 * parBDepth, parPlaneOfBranch)
        
        CP.Set dHlength / 2, 0, 0
        Set oHeaderTakeoffCurve1 = CreFlatOval(CP, 1.01 * parWidth, 1.01 * parDepth, parPlaneOfBranch)

        CP.Set -dHlength / 2, 0, 0
        Set oHeaderTakeoffCurve2 = CreFlatOval(CP, 1.01 * parWidth, 1.01 * parDepth, parPlaneOfBranch)

        CP.Set (dHlength / 2) - 0.004, 0, 0
        Set oHeaderStripCurve1 = CreFlatOval(CP, 1.06 * parWidth, 1.06 * parDepth, parPlaneOfBranch)

        CP.Set -(dHlength / 2) + 0.004, 0, 0
        Set oHeaderStripCurve2 = CreFlatOval(CP, 1.06 * parWidth, 1.06 * parDepth, parPlaneOfBranch)
        
    ElseIf parHVACShape = 4 Then 'Round=4
        CP.Set 0, dBranchStartY, 0
        Dir.Set 0, 1, 0
        Set oCurve1 = PlaceTrCircleByCenter(CP, Dir, dBranchDia / 2)
        
            
        CP.Set 0, dBLength, 0
        Set oCurve2 = PlaceTrCircleByCenter(CP, Dir, parBWidth / 2)
        
        Set oBranchTakeoffCurve = PlaceTrCircleByCenter(CP, Dir, 1.01 * parBWidth / 2)
        
        CP.Set 0, dBLength - 0.004, 0
        Set oBranchStripCurve = PlaceTrCircleByCenter(CP, Dir, 1.06 * parBWidth / 2)
        
        Dir.Set 1, 0, 0
        CP.Set dHlength / 2, 0, 0
        Set oHeaderTakeoffCurve1 = PlaceTrCircleByCenter(CP, Dir, 1.01 * parWidth / 2)
        
        CP.Set (dHlength / 2) - 0.004, 0, 0
        Set oHeaderStripCurve1 = PlaceTrCircleByCenter(CP, Dir, 1.06 * parWidth / 2)
        
        Dir.Set -1, 0, 0
        CP.Set -dHlength / 2, 0, 0
        Set oHeaderTakeoffCurve2 = PlaceTrCircleByCenter(CP, Dir, 1.01 * parWidth / 2)
        

        CP.Set -(dHlength / 2) + 0.004, 0, 0
        Set oHeaderStripCurve2 = PlaceTrCircleByCenter(CP, Dir, 1.06 * parWidth / 2)
        
        
        
    ElseIf parHVACShape = 1 Then
        
        If (parPlaneOfBranch = PI / 2) Then
            CP.Set 0, (parDepth / 2), 0
        Else
            CP.Set 0, (parWidth / 2), 0
        End If
        Set oCurve1 = CreRectBranch(CP, dBWidth, dBDepth, parPlaneOfBranch)
        
        CP.Set 0, dBLength, 0
            
        Set oCurve2 = CreRectBranch(CP, parBWidth, parBDepth, parPlaneOfBranch)
        
        Set oBranchTakeoffCurve = CreRectBranch(CP, 1.01 * parBWidth, 1.01 * parBDepth, parPlaneOfBranch)
        
        CP.Set 0, CP.y - 0.004, 0
        Set oBranchStripCurve = CreRectBranch(CP, 1.06 * parBWidth, 1.06 * parBDepth, parPlaneOfBranch)
        
        CP.Set dHlength / 2, 0, 0
        Set oHeaderTakeoffCurve1 = CreRectangle(CP, 1.01 * parWidth, 1.01 * parDepth, parPlaneOfBranch)

        CP.Set -dHlength / 2, 0, 0
        Set oHeaderTakeoffCurve2 = CreRectangle(CP, 1.01 * parWidth, 1.01 * parDepth, parPlaneOfBranch)

        CP.Set (dHlength / 2) - 0.004, 0, 0
        Set oHeaderStripCurve1 = CreRectangle(CP, 1.06 * parWidth, 1.06 * parDepth, parPlaneOfBranch)

        CP.Set -(dHlength / 2) + 0.004, 0, 0
        Set oHeaderStripCurve2 = CreRectangle(CP, 1.06 * parWidth, 1.06 * parDepth, parPlaneOfBranch)
        
    End If
    
'    If objProfile Is Nothing Then GoTo ErrorHandler
    
    Dim objBranchBody As Object
    Set objBranchBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oCurve1, oCurve2, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchBody
    
    Set objBranchBody = Nothing
    Set oCurve1 = Nothing
    Set oCurve2 = Nothing
    
' Place Code for Take-offs
    Dim objHeaderTakeoff1 As Object
    Dim objHeaderTakeoff2 As Object
    Dim objBranchTakeoff As Object
    Dim objHeaderStrip1 As Object
    Dim objHeaderStrip2 As Object
    Dim objBranchStrip As Object
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
    'Place Header Output 2(Take-off 1)
    oAxisVec.Set 1, 0, 0
    Set objHeaderTakeoff1 = PlaceProjection(m_OutputColl, oHeaderTakeoffCurve1, oAxisVec, 2 * INCH, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderTakeoff1
    Set objHeaderTakeoff1 = Nothing
    
    'Place Output 3 (Header Strip 1)
    Set objHeaderStrip1 = PlaceProjection(m_OutputColl, oHeaderStripCurve1, oAxisVec, 0.008, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderStrip1
    Set objHeaderStrip1 = Nothing
    
    'Place Output 4 (Header Take-off 2)
    oAxisVec.Set -1, 0, 0
    Set objHeaderTakeoff2 = PlaceProjection(m_OutputColl, oHeaderTakeoffCurve2, oAxisVec, 2 * INCH, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderTakeoff2
    Set objHeaderTakeoff2 = Nothing
    
    'Place Output 5(Header Strip 2)
    Set objHeaderStrip2 = PlaceProjection(m_OutputColl, oHeaderStripCurve2, oAxisVec, 0.008, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderStrip2
    Set objHeaderStrip2 = Nothing
    
    'Place Output 6 (Branch Take-off)
    oAxisVec.Set 0, 1, 0
    Set objBranchTakeoff = PlaceProjection(m_OutputColl, oBranchTakeoffCurve, oAxisVec, 2 * INCH, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchTakeoff
    Set objBranchTakeoff = Nothing
    
    'Place Output 7 (Branch Strip)
    Set objBranchStrip = PlaceProjection(m_OutputColl, oBranchStripCurve, oAxisVec, 0.008, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchStrip
    Set objBranchStrip = Nothing
    
    Set oHeaderStripCurve1 = Nothing
    Set oHeaderStripCurve2 = Nothing
    Set oHeaderTakeoffCurve1 = Nothing
    Set oHeaderTakeoffCurve2 = Nothing
    Set oBranchStripCurve = Nothing
    Set oBranchTakeoffCurve = Nothing
    
'Place Hvac Nozzle 1
    Dim oHvacNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle

    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim iPortIndex As Integer
       
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    Dim dNozzLength(1 To 3) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = dHlength
    dNozzLength(2) = 0.001
    dNozzLength(3) = 0.001
    
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
    Set NozzleFactory = Nothing
    
    CP.Set -dHlength / 2, 0, 0
    
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation CP
     
'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    
    PortDirection.Set -1, 0, 0
    iDistribPort.SetDirectionVector PortDirection

    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
        PortDirection.Set 0, 1, 0
    Else
        PortDirection.Set 0, 0, 1
    End If

    iDistribPort.SetRadialOrient PortDirection

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

' Place Nozzle 2
    PortDirection.Set 1, 0, 0
    CP.Set dHlength / 2, 0, 0
    iPortIndex = 2
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
     Set NozzleFactory = Nothing

     Set iDistribPort = oHvacNozzle
     iDistribPort.SetPortLocation CP

'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    iDistribPort.SetDirectionVector PortDirection
    
        If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
        PortDirection.Set 0, 1, 0
    Else
        PortDirection.Set 0, 0, 1
    End If
    iDistribPort.SetRadialOrient PortDirection
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
'Place Nozzle 3

   iPortIndex = 3
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), parBWidth, _
                                            parBDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim x, y, z As Double
    x = 0
    y = dBLength
    z = 0#
    CP.Set x, y, z

    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation CP

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    PortDirection.Set 0, 1, 0
    iDistribPort.SetDirectionVector PortDirection
    
    If CmpDblEqual(parPlaneOfBranch, PI / 2) Then

        PortDirection.Set 1, 0, 0
    Else

        PortDirection.Set 0, 0, 1
    End If
    iDistribPort.SetRadialOrient PortDirection
    
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CP = Nothing
    Set PortDirection = Nothing
    Set RadialDirection = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub
Public Function CreFlatOval(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x, CP.y - (Width - Depth) / 2, CP.z + Depth / 2
    Pt(2).Set CP.x, CP.y + (Width - Depth) / 2, Pt(1).z
    Pt(3).Set CP.x, CP.y + Width / 2, CP.z
    Pt(4).Set CP.x, Pt(2).y, CP.z - Depth / 2
    Pt(5).Set CP.x, Pt(1).y, Pt(4).z
    Pt(6).Set CP.x, CP.y - Width / 2, CP.z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 1, 0, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreFlatOval = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError METHOD
    
End Function
Public Function CreFlatOvalBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x - (Width - Depth) / 2, CP.y, CP.z + Depth / 2
    Pt(2).Set CP.x + (Width - Depth) / 2, CP.y, CP.z + Depth / 2
    Pt(3).Set CP.x + Width / 2, CP.y, CP.z
    Pt(4).Set CP.x + (Width - Depth) / 2, CP.y, CP.z - Depth / 2
    Pt(5).Set CP.x - (Width - Depth) / 2, CP.y, CP.z - Depth / 2
    Pt(6).Set CP.x - Width / 2, CP.y, CP.z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 0, 1, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreFlatOvalBranch = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError METHOD
    
End Function
Public Function CreRectangle(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreRectangle:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.x, CP.y - HW, CP.z + HD
    Pt(2).Set CP.x, CP.y + HW, CP.z + HD
    Pt(3).Set CP.x, CP.y + HW, CP.z - HD
    Pt(4).Set CP.x, CP.y - HW, CP.z - HD

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)
    Dim oDirVector As AutoMath.DVector
    
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 1, 0, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreRectangle = objCStr
    Set oLine = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    
    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError METHOD
   
End Function
Public Function CreRectBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreRectangle:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.x - HW, CP.y, CP.z + HD
    Pt(2).Set CP.x + HW, CP.y, CP.z + HD
    Pt(3).Set CP.x + HW, CP.y, CP.z - HD
    Pt(4).Set CP.x - HW, CP.y, CP.z - HD

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 0, 1, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreRectBranch = objCStr
    Set oLine = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    
    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError METHOD
   
End Function
'''<{(Circle begin)}>
Public Function PlaceTrCircleByCenter(ByRef centerPoint As AutoMath.DPosition, _
                            ByRef normalVector As AutoMath.DVector, _
                            ByRef Radius As Double) _
                            As IngrGeom3D.Circle3d

''' This function creates transient (non-persistent) circle
''' Example of call:
''' Dim point   As new AutoMath.DPosition
''' Dim normal  As new AutoMath.DVector
''' Dim objCircle  As IngrGeom3D.circle3d
''' point.set 0, 0, 0
''' normal.set 0, 0, 1
''' set objCircle = PlaceTrCircleByCenter(point, normal, 2 )
''' ......... use this object (e.g. to create projection)
''' set objCircle = Nothing


    Const METHOD = "PlaceTrCircleByCenter:"
    On Error GoTo ErrorHandler
        
    Dim oCircle As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the Circle"
    ' Create Circle object
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            centerPoint.x, centerPoint.y, centerPoint.z, _
                            normalVector.x, normalVector.y, normalVector.z, _
                            Radius)
    Set PlaceTrCircleByCenter = oCircle
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError METHOD

End Function

Private Sub ReportUnanticipatedError(InMethod As String)

Const METHOD = "ReportUnanticipatedError:"
'    Dim ern As IMSErrorService.JWellKnownErrorNumbers
'    Dim errorService As IMSErrorLog.IJErrorService
'    Dim oTrader As New Trader
'
'    Set errorService = oTrader.Service("ErrorHandler", "")
'
'    ern = errorService.ReportError(Err.Number, MODULE & ":" & InMethod, "UNANTICIPATED", App)
'
'    Select Case ern
'      Case imsAbortApplication:
'            errorService.TerminateApp
'    End Select
'
'    Set errorService = Nothing
'    Set oTrader = Nothing

    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub

