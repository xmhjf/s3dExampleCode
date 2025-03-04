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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Monday, June 25 2007
'   Description:
'   This symbol is prepared for Cotoured Flanged straight and conical tee taps of McGill Air flow corporation as per CR-120452
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CInsulation" 'Used for error messages

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
    Dim objInsTapBody As Object
    Dim oTopCurve As Object
    Dim objInsTakeOff As Object
    Dim objInsStrip As Object
    Dim oPlane As IngrGeom3D.Plane3d
    
    Dim dTapBodyWidth As Double
    Dim dTakeOffLoc As Double
    
    If parHVACShape = FlatOval Then
        Dim oFltOvalCurve As Object
        'Create the non persistent Duct (Header)
        oCenter.Set -(dBWidth / 2 + 10 * INCH), 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOval(oCenter, parWidth, _
                                                    parDepth, 0)
        Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    1, 0, 0, _
                                                    (dBWidth + 20 * INCH + 2 * parInsulationThickness), False)
        Set oFltOvalCurve = Nothing
        
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            oCenter.Set 0, 0, -parDepth / 2
            If CmpDblGreaterthan(dBDepth + 2 * parInsulationThickness, parDepth) Then
                oCenter.Set 0, -((parWidth / 2 - parDepth / 2) + (Sqr((parDepth / 2) ^ 2 - (dBDepth / 2) ^ 2))), -parDepth / 2
                Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                        dBDepth + 2 * parInsulationThickness, 0)
                oDirVector.Set 0, -1, 0
                Set objInsTapBody = PlaceProjection(m_OutputColl, oFltOvalCurve, _
                                    oDirVector, parWidth / 2 - Abs(oCenter.Y) + 1 * INCH, True)
            Else
                'Create the Non-Persistant TapBody
                Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                                        dBDepth + 2 * parInsulationThickness, 0)
                Set oTapBody = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                            0, -1, 0, _
                                                             parWidth / 2 + (5 * INCH) + parInsulationThickness, False)
                Set oFltOvalCurve = Nothing
                
                Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, _
                                                -(parWidth / 2 + (1 * INCH)), -parDepth / 2, 0, -1, 0)
            
                'Curve between TapBody and Header
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
                Set oIJcurve2 = oIntersectionCurve.Item(1)
                
                'Curve at the of the nozzle
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
                Set oIJcurve3 = oIntersectionCurve.Item(1)
                
                'Create Tap Body
                Set objInsTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oIJcurve2, oIJcurve3, True)
            End If
            dTakeOffLoc = 1 * INCH
            
        ElseIf lPartdatabasis = 30 Then
            Dim oFltOvalCurve1 As Object
            Dim oFltOvalCurve2 As Object
            Dim dTapBodyDepth As Double
            
            oCenter.Set 0, 0, -parDepth / 2
            dTapBodyDepth = dBDepth + 1 * INCH
            If CmpDblGreaterthan(dTapBodyDepth + 2 * parInsulationThickness, _
                                        parDepth) Then
                oCenter.Set 0, -((parWidth / 2 - parDepth / 2) + (Sqr((parDepth / 2) ^ 2 - (dBDepth / 2) ^ 2))), -parDepth / 2
                Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                        dBDepth + 2 * parInsulationThickness, 0)
                oDirVector.Set 0, -1, 0
                Set objInsTapBody = PlaceProjection(m_OutputColl, oFltOvalCurve, _
                                    oDirVector, parWidth / 2 - Abs(oCenter.Y) + 4 * INCH, True)
                
            Else
                'Create the Non-persistant TapBody
                Set oFltOvalCurve1 = CreFlatOvalBranch(oCenter, dBWidth + 1 * INCH + 2 * parInsulationThickness, _
                                                                dTapBodyDepth + 2 * parInsulationThickness, 0)
                oCenter.Set 0, -(parWidth / 2 + 4 * INCH), -parDepth / 2
                Set oFltOvalCurve2 = CreFlatOvalBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                                                dBDepth + 2 * parInsulationThickness, 0)
                
                Set oTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oFltOvalCurve1, _
                                                oFltOvalCurve2, True)
                
                'Create a Plane normal to the Branch to get the flat oval curve at the nozzle
                Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, _
                                                -(parWidth / 2 + (4 * INCH)), -parDepth / 2, 0, -1, 0)
            
                'Curve between TapBody and Header
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
                Set oIJcurve2 = oIntersectionCurve.Item(1)
                
                'Curve at the of the nozzle
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
                Set oIJcurve3 = oIntersectionCurve.Item(1)
                
                'Create Tap Body
                Set objInsTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oIJcurve2, oIJcurve3, True)
            End If
            dTakeOffLoc = 4 * INCH
        End If
            
        Set oFltOvalCurve1 = Nothing
        Set oFltOvalCurve2 = Nothing
        
        'Create the Take Off
        oCenter.Set 0, -(parWidth / 2 + dTakeOffLoc), -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, 1.01 * dBWidth + 2 * parInsulationThickness, _
                                                    dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness, 0)
        oDirVector.Set 0, -1, 0
        Set objInsTakeOff = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 2 * INCH, True)
        Set oFltOvalCurve = Nothing
        
        'Create Strip
        oCenter.Set 0, -(parWidth / 2 + dTakeOffLoc - 0.004), -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, 1.06 * dBWidth + 2 * parInsulationThickness, _
                                                dBDepth + 0.06 * dBDepth + 2 * parInsulationThickness, 0)
        Set objInsStrip = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 0.008, True)
        
    ElseIf parHVACShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            
            'Create the Tap Body
            oStPoint.Set -(dBWidth / 2 + parInsulationThickness), -(dBDepth / 2 + parInsulationThickness), 0
            oEnPoint.Set dBWidth / 2 + parInsulationThickness, _
                            dBDepth / 2 + parInsulationThickness, 1 * INCH
            Set objInsTapBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            dTakeOffLoc = 1 * INCH
            
        ElseIf lPartdatabasis = 30 Then
            
            'Create the Tap Body
            oStPoint.Set 0, 0, 0
            Set objInsTapBody = PlaceTrapezoid(m_OutputColl, oStPoint, dBWidth + 1 * INCH + 2 * parInsulationThickness, _
                                            dBDepth + 1 * INCH + 2 * parInsulationThickness, _
                                            dBWidth + 2 * parInsulationThickness, _
                                            dBDepth + 2 * parInsulationThickness, _
                                            4 * INCH, True, 0, 0, 0)
            dTakeOffLoc = 4 * INCH
        End If
        
        'Create Take Off
        oStPoint.Set -(1.01 * dBWidth / 2 + parInsulationThickness), _
                    -(dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness) / 2, dTakeOffLoc
        oEnPoint.Set (1.01 * dBWidth / 2 + parInsulationThickness), _
                        (dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness) / 2, dTakeOffLoc + 2 * INCH
        Set objInsTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Create Strip
        oStPoint.Set -(1.06 * dBWidth / 2 + parInsulationThickness), _
                        -(dBDepth + 0.06 * dBWidth + 2 * parInsulationThickness) / 2, dTakeOffLoc - 0.004
        oEnPoint.Set (1.06 * dBWidth / 2 + parInsulationThickness), _
                        (dBDepth + 0.06 * dBWidth + 2 * parInsulationThickness) / 2, dTakeOffLoc + 0.004
        Set objInsStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    ElseIf parHVACShape = 4 Then 'Round
        
        oStPoint.Set -(dBWidth + 10 * INCH + parInsulationThickness), 0, -parWidth / 2
        oEnPoint.Set (dBWidth + 10 * INCH + parInsulationThickness), 0, -parWidth / 2
        Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth + 2 * parInsulationThickness)
        
        If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
            
            If CmpDblGreaterthan(dBWidth + 2 * parInsulationThickness, parWidth) Then
                oStPoint.Set 0, 0, -(parWidth / 2 - (Sqr((parWidth / 2) ^ 2 - (dBWidth / 2) ^ 2)))
                oEnPoint.Set 0, 0, 1 * INCH
            
                Set objInsTapBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        dBWidth + 2 * parInsulationThickness, True)
            Else
                oStPoint.Set 0, 0, -parWidth / 2
                oEnPoint.Set 0, 0, 1 * INCH
                Set oTapBody = PlaceCylinderTrans(oStPoint, oEnPoint, dBWidth + 2 * parInsulationThickness)
                
                oCenter.Set 0, 0, 1 * INCH
                oDirVector.Set 0, 0, -1
                Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, (dBWidth + 2 * parInsulationThickness) / 2)
                
                oTransformationMat.LoadIdentity
                dRotAboutZaxis = PI
                oDirVector.Set 0, 0, 1
                oTransformationMat.Rotate dRotAboutZaxis, oDirVector
                oTopCurve.Transform oTransformationMat
            
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
                Set oIJcurve2 = oIntersectionCurve.Item(1)
                
                'Create Tap Body
                If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
                    Set objInsTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oTopCurve, oIJcurve2, True)
                ElseIf lPartdatabasis = 30 Then
                    Set objInsTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oIJcurve3, oIJcurve2, True)
                End If
            End If
            
            dTakeOffLoc = 1 * INCH
            
        ElseIf lPartdatabasis = 30 Then
            
            If CmpDblGreaterthan(dBWidth + 2 * INCH + 2 * parInsulationThickness, _
                                    parWidth + 2 * parInsulationThickness) Then
                
                oStPoint.Set 0, 0, -(parWidth / 2 - (Sqr((parWidth / 2) ^ 2 - (dBWidth / 2) ^ 2)))
                oEnPoint.Set 0, 0, 4 * INCH
                Set objInsTapBody = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                        dBWidth + 2 * parInsulationThickness + 1 * INCH, _
                                        dBWidth + 2 * parInsulationThickness, True)
            
            Else
                oStPoint.Set 0, 0, -(parWidth / 2 - (Sqr((parWidth / 2) ^ 2 - (dBWidth / 2) ^ 2)))
                oEnPoint.Set 0, 0, 4 * INCH
                Set oTapBody = PlaceConeTrans(oStPoint, oEnPoint, (dBWidth + 1 * INCH + 2 * parInsulationThickness) / 2, _
                                            (dBWidth + 2 * parInsulationThickness) / 2, True)
                
                Set oTopCurve = PlaceCylinderTrans(oStPoint, oEnPoint, dBWidth + 2 * parInsulationThickness)
                Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                0, 0, 4 * INCH, 0, 0, 1)
                Set oIJSurfaceHeader = oTopCurve
                oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
                Set oIJcurve3 = oIntersectionCurve.Item(1)
            
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
                Set oIJcurve2 = oIntersectionCurve.Item(1)
                
                'Create Tap Body
                If lPartdatabasis <= 1 Or lPartdatabasis = 25 Then
                    Set objInsTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oTopCurve, oIJcurve2, True)
                ElseIf lPartdatabasis = 30 Then
                    Set objInsTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oIJcurve3, oIJcurve2, True)
                End If
            End If
            dTakeOffLoc = 4 * INCH
            
        End If

        'Create Take Off
        oStPoint.Set 0, 0, dTakeOffLoc
        oEnPoint.Set 0, 0, dTakeOffLoc + 2 * INCH
        Set objInsTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                            1.01 * dBWidth + 2 * parInsulationThickness, True)
        
        'Create Strip
        oStPoint.Set 0, 0, dTakeOffLoc - 0.004
        oEnPoint.Set 0, 0, dTakeOffLoc + 0.004
        Set objInsStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    1.06 * dBWidth + 2 * parInsulationThickness, True)
    
    End If
    
    ' Set the output for Tap Body
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTapBody
    Set objInsTapBody = Nothing
    
    'Set the Output for TakeOff
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTakeOff
    Set objInsTakeOff = Nothing
    
    'Set the Output for Strip
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrip
    Set objInsStrip = Nothing
    
    'Remove the References
    Set oHeader = Nothing
    Set oTapBody = Nothing
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
    
Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub

End Sub


