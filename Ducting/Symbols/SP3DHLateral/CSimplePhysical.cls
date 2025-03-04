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
'   This symbol is created for Straight And Conical laterals of McGill Air Flow Corporation
'   The symbol supports the following two part data basis values
'   1.)Straight Lateral
'   2.)Conical Lateral
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CP As New AutoMath.DPosition
    
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    
    Dim Inch As Double
    Inch = 0.0254
    
    Dim PortDirection As AutoMath.DVector
    Set PortDirection = New AutoMath.DVector
    Dim RadialDirection As AutoMath.DVector
    Set RadialDirection = New AutoMath.DVector
    Dim CornerRadius As Double
    
    CornerRadius = 0
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parAngle = arrayOfInputs(6)
    parHVACShape = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
       
    iOutput = 0
        
    Dim dHlength As Double 'Header Length
    Dim dCentFrmPort2 As Double 'Distance from Center to Port 2
    Dim dBLength As Double 'Branch Length form center along the axis
    Dim dBranchInt2FromPort3 As Double 'Distance of intersection point 2 of branch(with header) and branch port along branch axis
    
    Dim dBottomBrWidth As Double 'Width of Branch at the Center(bottom)
    Dim dBottomBrDepth As Double 'Depth of Branch at the Center(bottom)

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
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
    
    Dim oBranchCurve1 As Object
    Dim oBranchCurve2 As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim objHeader As Object
    Dim objBranch As Object
    
    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector

'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    
'Checking for maximum values of Branch Width and Branch Depth based on part data basis
'--------------------------------------------
'    parWidth               Maximum BWidth
'--------------------------------------------
'    <=10inch               parWidth-3inch
'    >10inch and <=42inch   parWidth-4inch
'    >42inch                parWidth-5inch
    
    If (lPartdatabasis <= 1 Or lPartdatabasis = 45) Then
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            parBWidth = parWidth
        End If
        
        If (Not parHVACShape = 4) And CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
        
    ElseIf (lPartdatabasis = 50) Then
    
        If CmpDblLessThanOrEqualTo(parWidth, 10 * Inch) And CmpDblGreaterthan(parBWidth, parWidth - 3 * Inch) Then
            parBWidth = parWidth - (3 * Inch)
        ElseIf CmpDblGreaterthan(parWidth, 10 * Inch) And CmpDblLessThanOrEqualTo(parWidth, 42 * Inch) And CmpDblGreaterthan(parBWidth, parWidth - 4 * Inch) Then
            parBWidth = parWidth - (4 * Inch)
        ElseIf CmpDblGreaterthan(parWidth, 42 * Inch) And CmpDblGreaterthan(parBWidth, parWidth - 5 * Inch) Then
            parBWidth = parWidth - (5 * Inch)
        End If
        
        If (Not parHVACShape = 4) And CmpDblGreaterthan(parBDepth, parDepth - 2 * Inch) Then
            parBDepth = parDepth - 2 * Inch
        End If
    End If
    
    'In case of round shape (where Depth is optional) making Depth equal to Width
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    'In case of round shape (where BDepth is optional) making BDepth equal to BWidth
    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If
    
    
    Dim dAlpha As Double 'Half of the Difference between the branch width and width at intersection point of branch (with header)
    Dim dTemp As Double
       
'Calculations of header length, branch length and distance of port 2 from center based on McGill catalog
    If (lPartdatabasis <= 1 Or lPartdatabasis = 45) Then 'Straight Lateral
        dHlength = (parBWidth / Sin(parAngle)) + (2 * Inch)
        dAlpha = 0
        dBranchInt2FromPort3 = 2 * Inch
    ElseIf (lPartdatabasis = 50) Then 'Conical Lateral
        dHlength = ((parBWidth + (2 * Inch)) / Sin(parAngle)) + (2 * Inch)
        dTemp = ((parBWidth / Inch) + 2) / (4 * Tan(parAngle))
        dAlpha = 2 * Inch / (dTemp + 2)
        dBranchInt2FromPort3 = 4 * Inch
    End If
           
    dBLength = (parWidth / (2 * Sin(parAngle))) + _
              ((parBWidth + (2 * dAlpha)) / (2 * Tan(parAngle))) + dBranchInt2FromPort3
    dCentFrmPort2 = (parWidth / (2 * Tan(parAngle))) + _
              ((parBWidth + (2 * dAlpha)) / (2 * Sin(parAngle))) + Inch
    
'Assigning branch width and branch depth at the center(bottom) based on part data basis
    If (lPartdatabasis <= 1 Or lPartdatabasis = 45) Then 'Straight Lateral
         dBottomBrWidth = parBWidth
         dBottomBrDepth = parBDepth
    ElseIf (lPartdatabasis = 50) Then 'Conical Lateral
        dBottomBrWidth = parBWidth + ((dBLength * dAlpha) / (2 * Inch))
        dBottomBrDepth = parBDepth + (2 * Inch)
    End If
           
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
'Creating geometry based on shape of the lateral
    If parHVACShape = FlatOval Then
        Dim oHeaderCurve As Object
        
        'Creating flat oval header curve
        CP.Set -dHlength + dCentFrmPort2, 0, 0
        Set oHeaderCurve = CreFlatOval(CP, parWidth, parDepth, PI / 2)
        
        'Creating header
        oAxisVec.Set 1, 0, 0
        Set objHeader = PlaceProjection(m_OutputColl, oHeaderCurve, oAxisVec, dHlength, False)
        Set oHeaderCurve = Nothing
        
        'Creating branch curve at the center
        CP.Set 0, 0, 0
        Set oBranchCurve1 = CreFltOvlBranchNormaltoZ(CP, dBottomBrWidth, dBottomBrDepth, PI / 2 - (parAngle))
        
        'Creating branch curve at branch port
        CP.Set 0, 0, dBLength
        Set oBranchCurve2 = CreFltOvlBranchNormaltoZ(CP, parBWidth, parBDepth, PI / 2 - (parAngle))
        
        'Creating Branch by joining two branch curves
        Set objBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                             oBranchCurve1, oBranchCurve2, False)
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
        
        'Place Header Take-off 1
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
        
        'Place Header Take-off 2
        CP.Set dCentFrmPort2, 0, 0
        Set oHeaderTakeoffCurve2 = CreFlatOval(CP, 1.01 * parWidth, 1.01 * parDepth, PI / 2)
        
        oAxisVec.Set 1, 0, 0
        Set objHeaderTakeoff2 = PlaceProjection(m_OutputColl, oHeaderTakeoffCurve2, oAxisVec, 2 * Inch, True)
        Set oHeaderTakeoffCurve2 = Nothing
        
        'Creating Header Strip 2
        CP.Set dCentFrmPort2 - 0.004, 0, 0
        Set oHeaderStripCurve2 = CreFlatOval(CP, 1.025 * parWidth, 1.025 * parDepth, PI / 2)
        
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
        
    ElseIf parHVACShape = 4 Then 'Round=4
        
        'Creating Header
        oStPoint.Set -dHlength + dCentFrmPort2, 0, 0
        oEnPoint.Set dCentFrmPort2, 0, 0
        Set objHeader = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, False)
        
        'Check to see that branch width at center doesn't exceed width of the round duct
        If CmpDblGreaterthan(dBottomBrWidth, parWidth) Then
            dBottomBrWidth = parWidth
        End If
        
        'Creating Branch curve at the center
        CP.Set 0, 0, 0
        Dir.Set Cos(parAngle), 0, Sin(parAngle)
        Set oBranchCurve1 = PlaceTrCircleByCenter(CP, Dir, dBottomBrWidth / 2)
            
        'Creating Branch curve at the branch port
        CP.Set dBLength * Cos(parAngle), 0, dBLength * Sin(parAngle)
        Dir.Set Cos(parAngle), 0, Sin(parAngle)
        Set oBranchCurve2 = PlaceTrCircleByCenter(CP, Dir, parBWidth / 2)
        
        'Creating Branch by joining two branch curves
        Set objBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                             oBranchCurve1, oBranchCurve2, True)
        
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
        
        'Creating Header Take-off 1
        oStPoint.Set -dHlength + dCentFrmPort2, 0, 0
        oEnPoint.Set -dHlength + dCentFrmPort2 - (2 * Inch), 0, 0
        Set objHeaderTakeoff1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, True)
        
        'Creating Header Strip 1
        oStPoint.Set -dHlength + dCentFrmPort2 + 0.004, 0, 0
        oEnPoint.Set -dHlength + dCentFrmPort2 - 0.004, 0, 0
        Set objHeaderStrip1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.025 * parWidth, True)
        
        'Creating Header Take-off 2
        oStPoint.Set dCentFrmPort2, 0, 0
        oEnPoint.Set dCentFrmPort2 + (2 * Inch), 0, 0
        Set objHeaderTakeoff2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, True)
        
        'Creating Header Strip 2
        oStPoint.Set dCentFrmPort2 - 0.004, 0, 0
        oEnPoint.Set dCentFrmPort2 + 0.004, 0, 0
        Set objHeaderStrip2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.025 * parWidth, True)
    
        'Creating Branch Take-off
        oStPoint.Set dBLength * Cos(parAngle), 0, dBLength * Sin(parAngle)
        oEnPoint.Set (dBLength + 2 * Inch) * Cos(parAngle), 0, (dBLength + 2 * Inch) * Sin(parAngle)
        Set objBranchTakeoff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parBWidth, True)
        
        'Creating Branch Strip
        oStPoint.Set (dBLength - 0.004) * Cos(parAngle), 0, (dBLength - 0.004) * Sin(parAngle)
        oEnPoint.Set (dBLength + 0.004) * Cos(parAngle), 0, (dBLength + 0.004) * Sin(parAngle)
        Set objBranchStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.025 * parBWidth, True)
                
        
    ElseIf parHVACShape = Rectangular Then
        
        'Creating Header
        oStPoint.Set -dHlength + dCentFrmPort2, parDepth / 2, parWidth / 2
        oEnPoint.Set dCentFrmPort2, -parDepth / 2, -parWidth / 2
        Set objHeader = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
        'Creating branch curve at the center
        CP.Set dCentFrmPort2 - (dHlength / 2), 0, parWidth / 2
        Set oBranchCurve1 = CreRectBranchNormaltoZ(CP, dHlength - (2 * Inch), dBottomBrDepth, 0)
        
        'Creating branch curve at the branch port
        CP.Set 0, 0, dBLength
        Set oBranchCurve2 = CreRectBranchNormaltoZ(CP, parBWidth, parBDepth, PI / 2 - (parAngle))
        
        'Creating Branch by joining two branch curves
        Set objBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                             oBranchCurve1, oBranchCurve2, True)
                             
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
        
        'Creating Header Take-off 1
        oStPoint.Set -dHlength + dCentFrmPort2, 1.01 * parDepth / 2, 1.01 * parWidth / 2
        oEnPoint.Set -dHlength + dCentFrmPort2 - (2 * Inch), -1.01 * parDepth / 2, -1.01 * parWidth / 2
        Set objHeaderTakeoff1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Creating Header Strip 1
        oStPoint.Set -dHlength + dCentFrmPort2 + 0.004, 1.025 * parDepth / 2, 1.025 * parWidth / 2
        oEnPoint.Set -dHlength + dCentFrmPort2 - 0.004, -1.025 * parDepth / 2, -1.025 * parWidth / 2
        Set objHeaderStrip1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Creating Header Take-off 2
        oStPoint.Set dCentFrmPort2, 1.01 * parDepth / 2, 1.01 * parWidth / 2
        oEnPoint.Set dCentFrmPort2 + (2 * Inch), -1.01 * parDepth / 2, -1.01 * parWidth / 2
        Set objHeaderTakeoff2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Creating Header Strip 2
        oStPoint.Set dCentFrmPort2 - 0.004, 1.025 * parDepth / 2, 1.025 * parWidth / 2
        oEnPoint.Set dCentFrmPort2 + 0.004, -1.025 * parDepth / 2, -1.025 * parWidth / 2
        Set objHeaderStrip2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
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
        
    End If

    ' Set the Output 1(Header)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeader
    Set objHeader = Nothing
    
    ' Set the output 2(Branch)
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
    
'Creating Hvac Nozzle 1
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
    dNozzLength(1) = 0.001
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
    
    CP.Set -dHlength + dCentFrmPort2, 0, 0
    
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
    
    RadialDirection.Set 0, 1, 0
    iDistribPort.SetRadialOrient RadialDirection

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

' Place Nozzle 2
    PortDirection.Set 1, 0, 0
    CP.Set dCentFrmPort2, 0, 0
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

    RadialDirection.Set 0, 1, 0
    iDistribPort.SetRadialOrient RadialDirection
    
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
    CP.Set dBLength * Cos(parAngle), 0, dBLength * Sin(parAngle)

    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation CP

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    PortDirection.Set Cos(parAngle), 0, Sin(parAngle)
    iDistribPort.SetDirectionVector PortDirection

    RadialDirection.Set 0, 1, 0
    iDistribPort.SetRadialOrient RadialDirection
    
    
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
    ReportUnanticipatedError2 MODULE, METHOD
End Sub








