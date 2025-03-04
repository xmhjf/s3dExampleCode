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
'   ProgID:         SP3DHContFlgdLLTeeTap.HLLTeeTap
'   Author:         RUK
'   Creation Date:  Monday, June 25 2007
'   Description:
'   This symbol is prepared for Cotoured Flanged Lo-Loss tee taps of McGill Air flow corporation as per CR-120452
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
                                            
    If Not parHVACShape = 4 And Not parHVACShape = Rectangular And Not parHVACShape = FlatOval Then
        parHVACShape = 4
    End If
    
    Dim dBWidth As Double
    Dim dBDepth As Double
    
    dBWidth = parBWidth
    dBDepth = parBDepth
    
    If parHVACShape = 4 Then 'Round
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            dBWidth = parWidth
        End If
    ElseIf parHVACShape = FlatOval Then
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            dBDepth = parDepth
        End If
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            dBWidth = parWidth
        End If
    End If

    If parHVACShape = 4 Then 'round
        Dim dTapTopWidth As Double   'J
        If CmpDblLessThanOrEqualTo(dBWidth, parWidth - 2 * INCH) Then
            dTapTopWidth = dBWidth + 2 * INCH
        Else
            dTapTopWidth = dBWidth
        End If
    ElseIf parHVACShape = FlatOval Or parHVACShape = Rectangular Then
        Dim dTapTopDepth As Double   'J
        If CmpDblLessThanOrEqualTo(dBDepth, parDepth - 2 * INCH) Then
            dTapTopDepth = dBDepth + 2 * INCH
        Else
            dTapTopDepth = dBWidth
        End If
    End If
      
'        parBWidth (Round) or parBDepth (Rectangular)            dTapBodyHeight
'                  (Inches)                                        (Inches)
'                    3 - 8                                               4
'                    8.5 - 14                                            7
'                    14.5 - 26                                           10
'                    27 or Larger                                        13
    
    Dim dTapHeight As Double
    If CmpDblGreaterthan(dBWidth, 3 * INCH) And CmpDblLessThan(dBWidth, 8 * INCH) Then
        dTapHeight = 4 * INCH
    ElseIf CmpDblGreaterthan(dBWidth, 8.5 * INCH) And CmpDblLessThan(dBWidth, 14 * INCH) Then
        dTapHeight = 7 * INCH
    ElseIf CmpDblGreaterthan(dBWidth, 14.5 * INCH) And CmpDblLessThan(dBWidth, 26 * INCH) Then
        dTapHeight = 10 * INCH
    ElseIf CmpDblGreaterthan(dBWidth, 27 * INCH) Then
        dTapHeight = 13 * INCH
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
    
    Dim dTakeOffLoc As Double
    
    If parHVACShape = FlatOval Then
        Dim oFltOvalCurve As Object
        'Create the non persistent Duct (Header)
        oCenter.Set -(dBWidth + dTapHeight + 2 * INCH), 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOval(oCenter, parWidth, parDepth, 0)
        Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    1, 0, 0, _
                                                    (dBWidth + dTapHeight + 10 * INCH), False)
        
        'Create non persistant object to get the Bottom Curve of the TapBody
        oCenter.Set -(dTapHeight / 2), 0, -parDepth / 2
        oDirVector.Set 0, -1, 0
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, (dBWidth + dTapHeight), dTapTopDepth, 0)
        
        Set oTapBody = PlaceProjectionTrans(oFltOvalCurve, oDirVector, parWidth + dTapHeight, True)
        
        'Create non persistant object to get the Pad curve
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, (dBWidth + dTapHeight + 1 * INCH), dTapTopDepth + 1 * INCH, 0)
        Set oPad = PlaceProjectionTrans(oFltOvalCurve, oDirVector, parWidth + dTapHeight, True)
        
        'Create non persistant object to get the flat-oval shape at the nozzle
        oCenter.Set 0, 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth, dBDepth, 0)
        Set oTopCurve = PlaceProjectionTrans(oFltOvalCurve, oDirVector, (parWidth / 2 + dTapHeight), True)
        
        Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, -(parWidth / 2 + dTapHeight), -parDepth / 2, _
                                                0, -1, 0)
        
        dTakeOffLoc = dTapHeight
        
        'Curve between Pad and Header
        Set oIJSurfaceHeader = oPad
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve1 = oIntersectionCurve.Item(1)

        'Curve between TapBody and Header
        Set oIJSurfaceHeader = oTapBody
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve2 = oIntersectionCurve.Item(1)

        'Curve at the of the nozzle
        Set oIJSurfaceHeader = oTopCurve
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
        Set oFltOvalCurve = Nothing
        
    ElseIf parHVACShape = Rectangular Then
        'Create the Pad
        Dim oTopSurPts(0 To 3) As IJDPosition
        Dim oBotSurPts(0 To 3) As IJDPosition
        Dim iCount As Integer
        Dim oTapCollection As New Collection
        
        For iCount = 0 To 3
            Set oTopSurPts(iCount) = New DPosition
            Set oBotSurPts(iCount) = New DPosition
        Next iCount
        
        oTopSurPts(0).Set dBWidth / 2, dBDepth / 2, dTapHeight
        oTopSurPts(1).Set oTopSurPts(0).X, -oTopSurPts(0).Y, oTopSurPts(0).Z
        oTopSurPts(2).Set -oTopSurPts(0).X, -oTopSurPts(0).Y, oTopSurPts(0).Z
        oTopSurPts(3).Set -oTopSurPts(0).X, oTopSurPts(0).Y, oTopSurPts(0).Z
        
        oBotSurPts(0).Set dBWidth / 2, dTapTopDepth / 2, 0
        oBotSurPts(1).Set oBotSurPts(0).X, -oBotSurPts(0).Y, oBotSurPts(0).Z
        oBotSurPts(2).Set -(dBWidth / 2 + dTapHeight), -oBotSurPts(0).Y, oBotSurPts(0).Z
        oBotSurPts(3).Set oBotSurPts(2).X, oBotSurPts(0).Y, oBotSurPts(0).Z
        
        Set oTapCollection = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To oTapCollection.Count
            m_OutputColl.AddOutput "TapBody_", oTapCollection(iCount)
        Next iCount
        
        For iCount = 0 To 3
            Set oTopSurPts(iCount) = Nothing
            Set oBotSurPts(iCount) = Nothing
        Next iCount
        
        For iCount = 1 To oTapCollection.Count
            oTapCollection.Remove 1
        Next iCount
        
        'Create the Pad
        oStPoint.Set (dBWidth / 2 + INCH / 2), (dTapTopDepth / 2 + INCH / 2), 0
        oEnPoint.Set -(dBWidth / 2 + dTapHeight + INCH / 2), -(dTapTopDepth / 2 + INCH / 2), 0.001
        Set objPad = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        dTakeOffLoc = dTapHeight

        'Create Take Off
        oStPoint.Set -(1.01 * dBWidth / 2), -(dBDepth + 0.01 * dBWidth) / 2, dTakeOffLoc
        oEnPoint.Set (1.01 * dBWidth / 2), (dBDepth + 0.01 * dBWidth) / 2, dTakeOffLoc + 2 * INCH
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Create Strip
        oStPoint.Set -(1.06 * dBWidth / 2), -(dBDepth + 0.06 * dBWidth) / 2, dTakeOffLoc - 0.004
        oEnPoint.Set (1.06 * dBWidth / 2), (dBDepth + 0.06 * dBWidth) / 2, dTakeOffLoc + 0.004
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    ElseIf parHVACShape = 4 Then 'Round
        
        oStPoint.Set -(dBWidth + dTapHeight + 13 * INCH), 0, -parWidth / 2
        oEnPoint.Set (dBWidth + dTapHeight + 13 * INCH), 0, -parWidth / 2
        Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth)
        
        'Create the Tap Body
        Dim oEllipse As IngrGeom3D.Ellipse3d
        Dim oMajor As AutoMath.DPosition
        Dim dMMRatio As Double
        
        Set oMajor = New AutoMath.DPosition
        oCenter.Set -(dTapHeight / 2), 0, -parWidth / 2
        oMajor.Set (dBWidth + dTapHeight) / 2, 0, 0
        dMMRatio = dTapTopWidth / (dBWidth + dTapHeight)
        
        Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                oCenter.X, oCenter.Y, oCenter.Z, 0, 0, 1, _
                                oMajor.X, oMajor.Y, oMajor.Z, dMMRatio)
        oDirVector.Set 0, 0, 1
        Set oTapBody = PlaceProjectionTrans(oEllipse, oDirVector, parWidth / 2 + dTapHeight, True)
        
        'Create the pad
        oMajor.Set oMajor.X + INCH / 2, 0, 0
        dMMRatio = (dTapTopWidth + INCH) / (dBWidth + dTapHeight + INCH)
        
        Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                oCenter.X, oCenter.Y, oCenter.Z, 0, 0, 1, _
                                oMajor.X, oMajor.Y, oMajor.Z, dMMRatio)
        Set oPad = PlaceProjectionTrans(oEllipse, oDirVector, parWidth / 2 + dTapHeight, True)
        
        Set oEllipse = Nothing
        Set oMajor = Nothing
        
        'Create the Circle at the Nozzle
        Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, 0, dTapHeight, 0, 0, 1)
        
        oStPoint.Set 0, 0, -parDepth / 2
        oEnPoint.Set 0, 0, dTapHeight
        Set oTopCurve = PlaceCylinderTrans(oStPoint, oEnPoint, dBWidth)
        
        oTransformationMat.LoadIdentity
        dRotAboutZaxis = PI
        oDirVector.Set 0, 0, 1
        oTransformationMat.Rotate dRotAboutZaxis, oDirVector
        oTopCurve.Transform oTransformationMat
        
        dTakeOffLoc = dTapHeight
            
        Set oIJSurfaceHeader = oPad
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve1 = oIntersectionCurve.Item(1)
        
        Set oIJSurfaceHeader = oHeader
        oIJSurfaceHeader.Intersect oTapBody, oIntersectionCurve, IntersectCode
        Set oIJcurve2 = oIntersectionCurve.Item(1)
        
        Set oIJSurfaceHeader = oPlane
        oIJSurfaceHeader.Intersect oTopCurve, oIntersectionCurve, IntersectCode
        Set oIJcurve3 = oIntersectionCurve.Item(1)

        'Create the Pad
        Set objPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oIJcurve1, oIJcurve2, False)
        
        'Create Tap Body
        Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                    oIJcurve2, oIJcurve3, True)
            
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
    If parHVACShape = 4 Or parHVACShape = FlatOval Then
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "TapBody_", objTapBody
        Set objTapBody = Nothing
    End If
    
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
        Z = dTapHeight
    ElseIf parHVACShape = FlatOval Then
        X = 0
        Y = -(parWidth / 2 + dTapHeight)
        Z = -parDepth / 2
    ElseIf parHVACShape = Rectangular Then
        X = 0
        Y = 0
        Z = dTapHeight
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

