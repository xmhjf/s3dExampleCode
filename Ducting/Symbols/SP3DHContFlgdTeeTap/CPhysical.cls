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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DHContFlgdTeeTap.HTeeTap
'   Author:         RUK
'   Creation Date:  Thursday, June 21 2007
'   Description:
'   This symbol is prepared for Cotoured Flanged straight and conical tee taps of McGill Air flow corporation as per CR-120452
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CPhysical" 'Used for error messages

Private PI       As Double
Private Const INCH = 0.0254

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
      PI = 4 * Atn(1)
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
                Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHVACShape As Long
    Dim parInsulationThickness As Double
    
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHVACShape = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
            
    If CmpDblEqual(parDepth, LINEAR_TOLERANCE) Then
        parDepth = parWidth
    End If
    
    'Origin is taken at the top surface of the duct.
    'For this symbol orientation of axes are X-towards East and Z-along UP.
                                            
    'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    If Not parHVACShape = 4 And Not parHVACShape = Rectangular And Not parHVACShape = FlatOval Then
        parHVACShape = 4
    End If
    
    Dim dBWidth As Double
    Dim dBDepth As Double
    
    dBWidth = parBWidth
    dBDepth = parBDepth
    
    If parHVACShape = 4 Then 'Round
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            If CmpDblGreaterthan(parBWidth, parWidth) Then
                dBWidth = parWidth
            End If
        ElseIf lPartdatabasis = 30 Then
            If CmpDblGreaterthan(parBWidth, parWidth - 2 * INCH) Then
                dBWidth = parWidth - 2 * INCH
            End If
        End If
    ElseIf parHVACShape = FlatOval Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Or lPartdatabasis = 30 Then
            If CmpDblGreaterthan(parBDepth, parDepth) Then
                dBDepth = parDepth
            End If
            If CmpDblGreaterthan(parBWidth, parWidth) Then
                dBWidth = parWidth
            End If
        End If
    End If
    
    iOutput = 0
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint     As AutoMath.DPosition
    Dim oEnPoint     As AutoMath.DPosition
    Dim oCenter      As AutoMath.DPosition
    Dim oDirVector   As AutoMath.DVector
    Dim oTransformationMat  As AutoMath.DT4x4
    
    Dim oIJSurfaceHeader As IJSurface
    Dim oIntersectionCurve As IJElements
    Dim IntersectCode As Geom3dIntersectConstants
    Dim dRotAboutZaxis As Double
    
    Dim oIJcurve1 As IJCurve
    Dim oIJcurve2 As IJCurve
    Dim oIJcurve3 As IJCurve
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    Set oDirVector = New AutoMath.DVector
    Set oTransformationMat = New AutoMath.DT4x4
    
    Dim oHeader As Object
    Dim oTapBody  As Object
    Dim objTapBody As Object
    Dim oPad As Object
    Dim objPad As Object
    Dim oTopCurve As Object
    Dim objTakeOff As Object
    Dim objStrip As Object
    Dim oPlane As IngrGeom3D.Plane3d
    
    Dim dPadDepth As Double
    Dim dPadWidth As Double
    Dim dTapBodyWidth As Double
    Dim dPadBodyDepth As Double
    Dim dTakeOffLoc As Double
    
    If parHVACShape = FlatOval Then
        Dim oFltOvalCurve As Object
        'Create the non persistent Duct (Header)
        oCenter.Set -(dBWidth / 2 + 2 * INCH), 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOval(oCenter, parWidth, parDepth, 0)
        Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    1, 0, 0, _
                                                    (dBWidth + 10 * INCH), False)
        Set oFltOvalCurve = Nothing
        
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            'Create the Non-persistant Pad
            oCenter.Set 0, 0, -parDepth / 2
            dPadDepth = parBDepth + 1 * INCH
            If CmpDblGreaterthan(dPadDepth, parDepth) Then
                dPadDepth = parDepth - 0.001
            End If
            dPadWidth = dBWidth + 1 * INCH
            
            Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dPadWidth, dPadDepth, 0)
            Set oPad = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, 0, -1, 0, _
                                                     (parWidth / 2 + 5 * INCH), False)
            Set oFltOvalCurve = Nothing
            
            'Create the Non-Persistant TapBody
            oCenter.Set 0, 0, -parDepth / 2
            If CmpDblGreaterthan(dBDepth, parDepth - 0.002) Then
                dBDepth = parDepth - 0.002
            End If
            
            Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth, dBDepth, 0)
            Set oTapBody = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                        0, -1, 0, _
                                                         parWidth / 2 + (5 * INCH), False)
            Set oFltOvalCurve = Nothing
            
            Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, _
                                            -(parWidth / 2 + (1 * INCH)), -parDepth / 2, 0, -1, 0)
            dTakeOffLoc = 1 * INCH
            
        ElseIf lPartdatabasis = 30 Then
            Dim oFltOvalCurve1 As Object
            Dim oFltOvalCurve2 As Object
            Dim dTapBodyDepth As Double
            
            'Create the Non-persistant Pad
            oCenter.Set 0, 0, -parDepth / 2
            dPadDepth = dBDepth + 2 * INCH
            If CmpDblGreaterthan(dPadDepth, parDepth) Then
                dPadDepth = parDepth - 0.002
            End If
            Set oFltOvalCurve1 = CreFlatOvalBranch(oCenter, dBWidth + 2 * INCH, dPadDepth, 0)
            
            oCenter.Set 0, -(parWidth / 2 + 4 * INCH), -parDepth / 2
            Set oFltOvalCurve2 = CreFlatOvalBranch(oCenter, dBWidth + 1 * INCH, dPadDepth - 1 * INCH, 0)
            
            Set oPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oFltOvalCurve1, _
                                                    oFltOvalCurve2, True)
            
            'Create the Non-persistant TapBody
            oCenter.Set 0, 0, -parDepth / 2
            dTapBodyDepth = dBDepth + 1 * INCH
            If CmpDblGreaterthan(dTapBodyDepth, parDepth) Then
                dTapBodyDepth = parDepth - 0.002
            End If
            Set oFltOvalCurve1 = CreFlatOvalBranch(oCenter, dBWidth + 1 * INCH, _
                                                            dTapBodyDepth, 0)
            oCenter.Set 0, -(parWidth / 2 + 4 * INCH), -parDepth / 2
            Set oFltOvalCurve2 = CreFlatOvalBranch(oCenter, dBWidth, parBDepth, 0)
            
            Set oTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oFltOvalCurve1, _
                                            oFltOvalCurve2, True)
            
            'Create a Plane normal to the Branch to get the flat oval curve at the nozzle
            Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, _
                                            -(parWidth / 2 + (4 * INCH)), -parDepth / 2, 0, -1, 0)
            dTakeOffLoc = 4 * INCH
        End If
            
        Set oFltOvalCurve1 = Nothing
        Set oFltOvalCurve2 = Nothing
        
        'Curve between Pad and Header
        Set oIJSurfaceHeader = oPad
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve1 = oIntersectionCurve.Item(1)
        
        'Curve between TapBody and Header
        Set oIJSurfaceHeader = oTapBody
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve2 = oIntersectionCurve.Item(1)
        
        'Curve at the of the nozzle
        Set oIJSurfaceHeader = oTapBody
        oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
        Set oIJcurve3 = oIntersectionCurve.Item(1)
        
        'Create Pad
        Set objPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oIJcurve1, oIJcurve2, True)
        'Create Tap Body
        Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oIJcurve2, oIJcurve3, True)
        
        'Create the Take Off
        oCenter.Set 0, -(parWidth / 2 + dTakeOffLoc), -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, 1.01 * dBWidth, dBDepth + 0.01 * dBWidth, 0)
        oDirVector.Set 0, -1, 0
        Set objTakeOff = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 2 * INCH, True)
        Set oFltOvalCurve = Nothing
        
        'Create Strip
        oCenter.Set 0, -(parWidth / 2 + dTakeOffLoc - 0.004), -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, 1.06 * dBWidth, dBDepth + 0.06 * dBDepth, 0)
        Set objStrip = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 0.008, True)
        
    ElseIf parHVACShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            'Create the Pad
            oStPoint.Set -(dBWidth / 2 + INCH / 2), -(dBDepth / 2 + INCH / 2), 0
            oEnPoint.Set (dBWidth / 2 + INCH / 2), (dBDepth / 2 + INCH / 2), 0.001
            Set objPad = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Create the Tap Body
            oStPoint.Set -dBWidth / 2, -dBDepth / 2, 0
            oEnPoint.Set dBWidth / 2, dBDepth / 2, 1 * INCH
            Set objTapBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            dTakeOffLoc = 1 * INCH
            
        ElseIf lPartdatabasis = 30 Then
            
            'Create the Pad
            oStPoint.Set -(dBWidth / 2 + 2 * INCH / 2), -(dBDepth / 2 + 2 * INCH / 2), 0
            oEnPoint.Set (dBWidth / 2 + 2 * INCH / 2), (dBDepth / 2 + 2 * INCH / 2), 0.001
            Set objPad = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Create the Tap Body
            oStPoint.Set 0, 0, 0
            Set objTapBody = PlaceTrapezoid(m_OutputColl, oStPoint, dBWidth + 1 * INCH, dBDepth + 1 * INCH, _
                                            dBWidth, dBDepth, _
                                            4 * INCH, True, 0, 0, 0)
            dTakeOffLoc = 4 * INCH
        End If
        
        'Create Take Off
        oStPoint.Set -(1.01 * dBWidth / 2), -(dBDepth + 0.01 * dBWidth) / 2, dTakeOffLoc
        oEnPoint.Set (1.01 * dBWidth / 2), (dBDepth + 0.01 * dBWidth) / 2, dTakeOffLoc + 2 * INCH
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Create Strip
        oStPoint.Set -(1.06 * dBWidth / 2), -(dBDepth + 0.06 * dBWidth) / 2, dTakeOffLoc - 0.004
        oEnPoint.Set (1.06 * dBWidth / 2), (dBDepth + 0.06 * dBWidth) / 2, dTakeOffLoc + 0.004
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    ElseIf parHVACShape = 4 Then 'Round
        
        oStPoint.Set -(dBWidth + 10 * INCH), 0, -parWidth / 2
        oEnPoint.Set (dBWidth + 10 * INCH), 0, -parWidth / 2
        Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth)
        
        Dim dAngle1 As Double
        Dim dAngle2 As Double
        Dim dTotAng As Double
        
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            
            If CmpDblEqual(parWidth, dBWidth) Then
                dAngle1 = PI / 2
                dAngle2 = 0
            Else
                dAngle1 = Atn(dBWidth / Sqr(parWidth ^ 2 - dBWidth ^ 2))
                dAngle2 = INCH / parWidth
            End If
            dTotAng = dAngle1 + dAngle2
            
            oStPoint.Set 0, 0, -parWidth / 2
            oEnPoint.Set 0, 0, 1 * INCH
            Set oPad = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth * Sin(dTotAng))
        
            oStPoint.Set 0, 0, -parWidth / 2
            oEnPoint.Set 0, 0, 1 * INCH
            Set oTapBody = PlaceCylinderTrans(oStPoint, oEnPoint, dBWidth)
            
            oCenter.Set 0, 0, 1 * INCH
            oDirVector.Set 0, 0, -1
            Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, dBWidth / 2)
            
            oTransformationMat.LoadIdentity
            dRotAboutZaxis = PI
            oDirVector.Set 0, 0, 1
            oTransformationMat.Rotate dRotAboutZaxis, oDirVector
            oTopCurve.Transform oTransformationMat
            
            dTakeOffLoc = 1 * INCH
            
        ElseIf lPartdatabasis = 30 Then
            
            oStPoint.Set 0, 0, -parWidth / 2    'Base Center
            oEnPoint.Set 0, 0, 4 * INCH         'Top Center
            If CmpDblGreaterthan(dBWidth + 2 * INCH, parWidth) Then
                dBWidth = parWidth - 2 * INCH - 0.002
            End If
            Set oPad = PlaceConeTrans(oStPoint, oEnPoint, (dBWidth + 2 * INCH) / 2, (dBWidth + 1 * INCH) / 2, True)
            
            Set oTapBody = PlaceConeTrans(oStPoint, oEnPoint, (dBWidth + 1 * INCH) / 2, dBWidth / 2, True)
            
            Set oTopCurve = PlaceCylinderTrans(oStPoint, oEnPoint, dBWidth)
            Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                            0, 0, 4 * INCH, 0, 0, 1)
            Set oIJSurfaceHeader = oTopCurve
            oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
            Set oIJcurve3 = oIntersectionCurve.Item(1)
            
            dTakeOffLoc = 4 * INCH
            
        End If

            Set oIJSurfaceHeader = oPad
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve1 = oIntersectionCurve.Item(1)
            
            Set oIJSurfaceHeader = oTapBody
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve2 = oIntersectionCurve.Item(1)
            
            'Create the Pad
            Set objPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oIJcurve1, oIJcurve2, False)
            
            'Create Tap Body
            If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
                Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oTopCurve, oIJcurve2, True)
            ElseIf lPartdatabasis = 30 Then
                Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oIJcurve3, oIJcurve2, True)
            End If
            
            'Create Take Off
            oStPoint.Set 0, 0, dTakeOffLoc
            oEnPoint.Set 0, 0, dTakeOffLoc + 2 * INCH
            Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * dBWidth, True)
            
            'Create Strip
            oStPoint.Set 0, 0, dTakeOffLoc - 0.004
            oEnPoint.Set 0, 0, dTakeOffLoc + 0.004
            Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * dBWidth, True)
    
    End If
    
    'Set the Output for Pad
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPad
    Set objPad = Nothing
    
    ' Set the output for Tap Body
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTapBody
    Set objTapBody = Nothing
    
    'Set the Output for TakeOff
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing
    
    'Set the Output for Strip
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
    'Remove the References
    Set oHeader = Nothing
    Set oTapBody = Nothing
    Set oPad = Nothing
    Set oTopCurve = Nothing
    Set oPlane = Nothing
    Set oIJSurfaceHeader = Nothing
    Set oIntersectionCurve = Nothing
    
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oDirVector = Nothing
    Set oTransformationMat = Nothing
    Set oIJcurve1 = Nothing
    Set oIJcurve2 = Nothing
    Set oIJcurve3 = Nothing
    
'-------------------------------------------------------------------------------------------------
'==================================================
'BUILD HVACNOZZLE1 at the begining of the Header
'==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim oDir As New AutoMath.DVector
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    
    Dim dWidth As Double
    Dim dDepth As Double
    Dim CornerRadius As Double
    Dim lCSType As Long
    
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
    
    'Set HVAC nozzle parameters
    iPortIndex = 1
    
    If parHVACShape = Rectangular Then
        dDepth = dBDepth
        dWidth = dBWidth
        lCSType = GSCADNozzleEntities.Rectangular
    ElseIf parHVACShape = 4 Then 'Round
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
        dDepth = dBWidth
        dWidth = dBWidth
        lCSType = GSCADNozzleEntities.Round
    ElseIf parHVACShape = FlatOval Then
        dDepth = dBDepth
        dWidth = dBWidth
        lCSType = GSCADNozzleEntities.FlatOval
    End If
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HvacNozzle1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)

    
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    If parHVACShape = 4 Then 'Round
        X = 0
        Y = 0
        Z = 1 * INCH
        If lPartdatabasis = 30 Then
            Z = 4 * INCH
        End If
    ElseIf parHVACShape = FlatOval Then
        X = 0
        Y = -(parWidth / 2 + 1 * INCH)
        If lPartdatabasis = 30 Then
            Y = -(parWidth / 2 + 4 * INCH)
        End If
        Z = -parDepth / 2
    ElseIf parHVACShape = Rectangular Then
        X = 0
        Y = 0
        Z = 1 * INCH
        If lPartdatabasis = 30 Then
            Z = 4 * INCH
        End If
    End If
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 0, 0, 1
    oDistribPort.SetDirectionVector dir
    If parHVACShape = FlatOval Then
        dir.Set 0, -1, 0
        oDistribPort.SetDirectionVector dir
        oDir.Set 0, 0, 1
        oDistribPort.SetRadialOrient oDir
    ElseIf parHVACShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set NozzleFactory = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub

End Sub

