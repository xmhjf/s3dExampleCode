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
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   ProgID:         SP3DHTaperBodyLateral.HTapLateral
'   Author:         VRK
'   Creation Date:  Thursday,June 28 2007
'   Description:
'   This symbol is prepared for Tapered Body Lateral  of McGill Air flow corporation as per CR-120452
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private Const E_FAIL = &H80004005
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI  As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput  As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim PortDirection As New AutoMath.DVector
    Dim RadialDirection As New AutoMath.DVector
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHVACShape As Long
    Dim parWidth2 As Double
    Dim parDepth2 As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double
    
    Dim Inch As Double
    Inch = 0.0254
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHVACShape = arrayOfInputs(6)
    parWidth2 = arrayOfInputs(7)
    parDepth2 = arrayOfInputs(8)
    parAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
            
    'In case of round shape (where Depth is optional) making Depth equal to Width
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    'In case of round shape (where Depth2 is optional) making Depth2 equal to Width2
    If CmpDblEqual(parDepth2, 0) Then
        parDepth2 = parWidth2
    End If
    
    'In case of round shape (where BDepth is optional) making BDepth equal to BWidth
    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If
    
    If CmpDblGreaterthan(parBDepth, parDepth2) Then
        parBDepth = parDepth2
    End If

    Dim dHlength As Double 'Header Length
    Dim dCentFrmPort2 As Double 'Distance from Center to Port 2
    Dim dBLength As Double 'Branch Length form center along the axis
    Dim dTemp As Double 'Variable used for temporary calculations
    
    'The following calculations are made as per McGill catalog
    dTemp = (parWidth - parWidth2) / ((2 * Tan(parAngle)))
    dHlength = 2 * Inch + (parBWidth / (Sin(parAngle))) - dTemp
    dCentFrmPort2 = 1 * Inch + (parBWidth / ((2 * Sin(parAngle)))) + (parWidth2 / (2 * Tan(parAngle))) + (Inch * dTemp / dHlength)
    dBLength = 2 * Inch + (Inch * dTemp / dHlength) + (parWidth2 / (2 * Sin(parAngle))) + (parBWidth / (2 * Tan(parAngle)))


    Dim stPoint  As New AutoMath.DPosition
    Dim enPoint  As New AutoMath.DPosition
    
    Dim objBranch As Object
    Dim objTaperedBody As Object
    
    Dim objHeaderTakeoff1 As Object
    Dim objHeaderTakeoff2 As Object
    Dim objBranchTakeoff As Object
    Dim objHeaderStrip1 As Object
    Dim objHeaderStrip2 As Object
    Dim objBranchStrip As Object
    
    Dim oHeaderStripCurve1 As Object
    Dim oHeaderStripCurve2 As Object
    Dim oHeaderTakeoffCurve1 As Object
    Dim oHeaderTakeoffCurve2 As Object
    Dim oBranchStripCurve As Object
    Dim oBranchTakeoffCurve As Object
    
    Dim oHeaderCurve1 As Object
    Dim oHeaderCurve2 As Object
    Dim oBranchCurve1 As Object
    Dim oBranchCurve2 As Object
        
    Dim CP As AutoMath.DPosition
    Set CP = New AutoMath.DPosition
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
    If parHVACShape = 4 Then 'Round
    
        'Creating Tapered header
        stPoint.Set (dCentFrmPort2 - dHlength), 0, 0
        enPoint.Set dCentFrmPort2, 0, 0
        Set objTaperedBody = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, parWidth2 / 2, True)
        
        'Creating Branch
        stPoint.Set 0, 0, 0
        enPoint.Set dBLength * Cos(parAngle), 0, dBLength * Sin(parAngle)
        Set objBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.01 * parBWidth, True)
        
        'Creating Header Take-off 1
        stPoint.Set (dCentFrmPort2 - dHlength) - 2 * Inch, 0, 0
        enPoint.Set (dCentFrmPort2 - dHlength), 0, 0
        Set objHeaderTakeoff1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.01 * parWidth, True)
        
        'Creating Header Strip 1
        Dim objStrip As Object
        stPoint.Set (dCentFrmPort2 - dHlength) - 0.004, 0, 0
        enPoint.Set stPoint.x + 0.008, 0, 0
        Set objHeaderStrip1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.025 * parWidth, True)
        
        'Creating Header Take-off 2
        stPoint.Set dCentFrmPort2, 0, 0
        enPoint.Set dCentFrmPort2 + 2 * Inch, 0, 0
        Set objHeaderTakeoff2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.01 * parWidth2, True)
        
        'Creating Header Strip 2
        stPoint.Set dCentFrmPort2 - 0.004, 0, 0
        enPoint.Set stPoint.x + 0.008, 0, 0
        Set objHeaderStrip2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.025 * parWidth2, True)
        
  
        'Creating Branch Take off
        stPoint.Set dBLength * Cos(parAngle), 0, dBLength * Sin(parAngle)
        enPoint.Set (dBLength + (2 * Inch)) * Cos(parAngle), 0, (dBLength + (2 * Inch)) * Sin(parAngle)
        Set objBranchTakeoff = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.01 * parBWidth, True)
        
        'Creating Branch Strip
        stPoint.Set (dBLength - 0.004) * Cos(parAngle), 0, (dBLength - 0.004) * Sin(parAngle)
        enPoint.Set (dBLength + 0.004) * Cos(parAngle), 0, (dBLength + 0.004) * Sin(parAngle)
        Set objBranchStrip = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.025 * parBWidth, True)
    
    ElseIf parHVACShape = 1 Then
     
        'Creating Tapered header
        CP.Set dCentFrmPort2 - dHlength, 0, 0
        Set oHeaderCurve1 = CreRectangle(CP, parWidth, parDepth, PI / 2)
        
        CP.Set dCentFrmPort2, 0, 0
        Set oHeaderCurve2 = CreRectangle(CP, parWidth2, parDepth2, PI / 2)
        
        Set objTaperedBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oHeaderCurve1, oHeaderCurve2, True)
        
        'Creating Branch
        CP.Set 0, 0, 0
        Set oBranchCurve1 = CreRectBranchNormaltoZ(CP, parBWidth, parBDepth, PI / 2 - parAngle)
        
        CP.Set 0, 0, dBLength
        Set oBranchCurve2 = CreRectBranchNormaltoZ(CP, parBWidth, parBDepth, PI / 2 - parAngle)
        
        Set objBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oBranchCurve1, oBranchCurve2, True)
        
        'Creating Header Take-off 1
        stPoint.Set dCentFrmPort2 - dHlength - 2 * Inch, -1.01 * parDepth / 2, 1.01 * parWidth / 2
        enPoint.Set dCentFrmPort2 - dHlength, 1.01 * parDepth / 2, -1.01 * parWidth / 2
        Set objHeaderTakeoff1 = PlaceBox(m_OutputColl, stPoint, enPoint)
        
        'Creating Header Strip 1
        stPoint.Set dCentFrmPort2 - dHlength - 0.004, -1.025 * parDepth / 2, 1.025 * parWidth / 2
        enPoint.Set stPoint.x + 0.008, 1.025 * parDepth / 2, -1.025 * parWidth / 2
        Set objHeaderStrip1 = PlaceBox(m_OutputColl, stPoint, enPoint)
        
        'Creating Header Take-off 2
        stPoint.Set dCentFrmPort2, -1.01 * parDepth2 / 2, 1.01 * parWidth2 / 2
        enPoint.Set dCentFrmPort2 + 2 * Inch, 1.01 * parDepth2 / 2, -1.01 * parWidth2 / 2
        Set objHeaderTakeoff2 = PlaceBox(m_OutputColl, stPoint, enPoint)
        
        'Creating Header Strip 2
        stPoint.Set dCentFrmPort2 - 0.004, -1.025 * parDepth2 / 2, 1.025 * parWidth2 / 2
        enPoint.Set stPoint.x + 0.008, 1.025 * parDepth2 / 2, -1.025 * parWidth2 / 2
        Set objHeaderStrip2 = PlaceBox(m_OutputColl, stPoint, enPoint)
        
        'Creating Branch Take-off
        CP.Set 0, 0, dBLength
        Set oBranchTakeoffCurve = CreRectBranchNormaltoZ(CP, 1.01 * parBWidth, 1.01 * parBDepth, PI / 2 - (parAngle))
        
        oAxisVec.Set Cos(parAngle), 0, Sin(parAngle)
        Set objBranchTakeoff = PlaceProjection(m_OutputColl, oBranchTakeoffCurve, oAxisVec, 2 * Inch, True)
        Set oBranchTakeoffCurve = Nothing
        
        'Creating Branch Strip
        CP.Set 0, 0, dBLength - 0.004
        Set oBranchStripCurve = CreRectBranchNormaltoZ(CP, 1.025 * parBWidth, 1.025 * parBDepth, PI / 2 - (parAngle))
        
        Set objBranchStrip = PlaceProjection(m_OutputColl, oBranchStripCurve, oAxisVec, 0.008, True)
        Set oBranchStripCurve = Nothing
        
    ElseIf parHVACShape = FlatOval Then
        
        'Creating Tapered header
        CP.Set dCentFrmPort2 - dHlength, 0, 0
        Set oHeaderCurve1 = CreFlatOval(CP, parWidth, parDepth, PI / 2)
        
        CP.Set dCentFrmPort2, 0, 0
        Set oHeaderCurve2 = CreFlatOval(CP, parWidth2, parDepth2, PI / 2)

        Set objTaperedBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oHeaderCurve1, oHeaderCurve2, True)
        
        'Creating Branch
        CP.Set 0, 0, 0
        Set oBranchCurve1 = CreFltOvlBranchNormaltoZ(CP, parBWidth, parBDepth, PI / 2 - parAngle)
        
        CP.Set 0, 0, dBLength
        Set oBranchCurve2 = CreFltOvlBranchNormaltoZ(CP, parBWidth, parBDepth, PI / 2 - parAngle)
        
        Set objBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oBranchCurve1, oBranchCurve2, True)
        

        'Creating Header Take-off 1
        CP.Set -dHlength + dCentFrmPort2, 0, 0
        Set oHeaderTakeoffCurve1 = CreFlatOval(CP, 1.01 * parWidth, 1.01 * parDepth, PI / 2)
        
        oAxisVec.Set -1, 0, 0
        Set objHeaderTakeoff1 = PlaceProjection(m_OutputColl, oHeaderTakeoffCurve1, oAxisVec, 2 * Inch, True)
        Set oHeaderTakeoffCurve1 = Nothing
    
        'Creating Header Strip 1
        CP.Set -dHlength + dCentFrmPort2 + 0.004, 0, 0
        Set oHeaderStripCurve1 = CreFlatOval(CP, 1.025 * parWidth, 1.025 * parDepth, PI / 2)
        
        Set objHeaderStrip1 = PlaceProjection(m_OutputColl, oHeaderStripCurve1, oAxisVec, 0.008, True)
        Set oHeaderStripCurve1 = Nothing
        
        'Creating Header Take-off 2
        CP.Set dCentFrmPort2, 0, 0
        Set oHeaderTakeoffCurve2 = CreFlatOval(CP, 1.01 * parWidth2, 1.01 * parDepth2, PI / 2)
        
        oAxisVec.Set 1, 0, 0
        Set objHeaderTakeoff2 = PlaceProjection(m_OutputColl, oHeaderTakeoffCurve2, oAxisVec, 2 * Inch, True)
        Set oHeaderTakeoffCurve2 = Nothing
        
        'Creating Header Strip 2
        CP.Set dCentFrmPort2 - 0.004, 0, 0
        Set oHeaderStripCurve2 = CreFlatOval(CP, 1.025 * parWidth2, 1.025 * parDepth2, PI / 2)
        
        Set objHeaderStrip2 = PlaceProjection(m_OutputColl, oHeaderStripCurve2, oAxisVec, 0.008, True)
        Set oHeaderStripCurve2 = Nothing
    
        'Creating Branch Take-off
        CP.Set 0, 0, dBLength
        Set oBranchTakeoffCurve = CreFltOvlBranchNormaltoZ(CP, 1.01 * parBWidth, 1.01 * parBDepth, PI / 2 - (parAngle))
        
        oAxisVec.Set Cos(parAngle), 0, Sin(parAngle)
        Set objBranchTakeoff = PlaceProjection(m_OutputColl, oBranchTakeoffCurve, oAxisVec, 2 * Inch, True)
        Set oBranchTakeoffCurve = Nothing
        
        'Creating Branch Strip
        CP.Set 0, 0, dBLength - 0.004
        Set oBranchStripCurve = CreFltOvlBranchNormaltoZ(CP, 1.025 * parBWidth, 1.025 * parBDepth, PI / 2 - (parAngle))
        
        Set objBranchStrip = PlaceProjection(m_OutputColl, oBranchStripCurve, oAxisVec, 0.008, True)
        Set oBranchStripCurve = Nothing
    
    End If
    
    'Set output 1(Tapered Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTaperedBody
    Set objTaperedBody = Nothing
    
    'Set the Output 2(Branch)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranch
    Set objBranch = Nothing
    
    'Place Output 3 (Header Take off 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderTakeoff1
    Set objHeaderTakeoff1 = Nothing
    
    'Place Output 4 (Header Strip 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderStrip1
    Set objHeaderStrip1 = Nothing
    
    'Place Output 5(Header Take off 2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderTakeoff2
    Set objHeaderTakeoff2 = Nothing
        
    'Place Output 6 (Header Strip 2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderStrip2
    Set objHeaderStrip2 = Nothing
    
    'Place Output 7 (Branch Take-off)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchTakeoff
    Set objBranchTakeoff = Nothing
    
    'Place Output 8 (Branch Strip)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchStrip
    Set objBranchStrip = Nothing
    
    
    
    
'-------------------------------------------------------------------------------------------------
'==================================================
'BUILD HVACNOZZLE1 at the begining of the Header
'==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim pos As New AutoMath.DPosition
    Dim Dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector
    
    Dim CornerRadius As Double

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
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    pos.Set dCentFrmPort2 - dHlength, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Dir.Set -1, 0, 0
    oDistribPort.SetDirectionVector Dir
    
    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

'==================================================
'BUILD HVACNOZZLE2 at the end of the Reducing Tee
'==================================================
    iPortIndex = 2
    

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth2, _
                                            parDepth2, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
    
    'Position of the nozzle should be the conenct point of the nozzle
    pos.Set dCentFrmPort2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Dir.Set 1, 0, 0
    oDistribPort.SetDirectionVector Dir
    
    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
'============================================
'BUILD HVACNOZZLE3 at the end of the Branch
'============================================
    iPortIndex = 3

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", parHVACShape, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), parBWidth, _
                                            parBDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle

    pos.Set dBLength * Cos(parAngle), 0, dBLength * Sin(parAngle)
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Dir.Set Cos(parAngle), 0, Sin(parAngle)
    oDistribPort.SetDirectionVector Dir

    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oRadialDir = Nothing
    Set pos = Nothing
    Set Dir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub




