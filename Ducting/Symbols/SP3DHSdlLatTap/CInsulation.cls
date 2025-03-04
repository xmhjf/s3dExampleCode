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
'   Author:         VRK
'   Creation Date:  Wednesday, June 27 2007
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
    Dim parAngle As Double
    Dim parHVACShape As Long
    Dim parInsulationThickness As Double
    
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parAngle = arrayOfInputs(6)
    parHVACShape = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
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
    Dim dTapLength As Double
    Dim dAlpha As Double
    
    dBWidth = parBWidth
    dBDepth = parBDepth
    
    If parHVACShape = 4 Then 'Round
        If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
            If CmpDblGreaterthan(parBWidth, parWidth) Then
                dBWidth = parWidth - 0.002
            End If
        ElseIf lPartdatabasis = 80 Then
            If CmpDblLessThanOrEqualTo(parWidth, 10 * INCH) Then
                If CmpDblGreaterthan(parBWidth, parWidth - 3 * INCH) Then
                    dBWidth = parWidth - 3 * INCH
                End If
            ElseIf CmpDblLessThanOrEqualTo(parWidth, 42 * INCH) Then
                If CmpDblGreaterthan(parBWidth, parWidth - 4 * INCH) Then
                    dBWidth = parWidth - 4 * INCH
                End If
            Else
                If CmpDblGreaterthan(parBWidth, parWidth - 5 * INCH) Then
                    dBWidth = parWidth - 5 * INCH
                End If
            End If
        End If
    ElseIf parHVACShape = FlatOval Or parHVACShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
            If CmpDblGreaterthanOrEqualTo(parBDepth, parDepth) Then
                dBDepth = parDepth - 0.002
            End If
            If CmpDblGreaterthanOrEqualTo(parBWidth, parWidth) Then
                dBWidth = parWidth - 0.002
            End If
        
        ElseIf lPartdatabasis = 80 Then
            If CmpDblLessThanOrEqualTo(parWidth, 10 * INCH) Then
                If CmpDblGreaterthan(parBWidth, parWidth - 3 * INCH) Then
                    dBWidth = parWidth - 3 * INCH
                End If
            ElseIf CmpDblLessThanOrEqualTo(parWidth, 42 * INCH) Then
                If CmpDblGreaterthan(parBWidth, parWidth - 4 * INCH) Then
                    dBWidth = parWidth - 4 * INCH
                End If
            Else
                If CmpDblGreaterthan(parBWidth, parWidth - 5 * INCH) Then
                    dBWidth = parWidth - 5 * INCH
                End If
            End If
            If CmpDblGreaterthan(parBDepth, parDepth - 2 * INCH) Then
                dBDepth = parDepth - 2 * INCH
            End If
        End If
    End If
    
    'Caliculate the Length of the Tap as per McGill HVAC Catalog
    If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
        dTapLength = ((parWidth) / (2 * Sin(parAngle))) + _
                        ((dBWidth) / (2 * Tan(parAngle))) + 2 * INCH
    ElseIf lPartdatabasis = 80 Then
        If CmpDblLessThan(parAngle, PI / 6) Then
            parAngle = PI / 6
        End If
        Dim dTempApha As Double     'Temparary variable to caliculate Alpha
        Dim dTapLenFrmOrg As Double
    
        dTempApha = ((parBWidth / INCH) + 2) / (4 * Tan(parAngle))
        dAlpha = 2 * INCH / (dTempApha + 2)
        dTapLength = (parWidth / (2 * Sin(parAngle))) + _
              ((dBWidth + (2 * dAlpha)) / (2 * Tan(parAngle))) + 4 * INCH
    End If
    dTapLenFrmOrg = dTapLength - (parWidth / (2 * Sin(parAngle)))
    
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
    Dim oBotCurve As Object
    Dim objTakeOff As Object
    Dim objStrip As Object
    Dim oPlane As IngrGeom3D.Plane3d
    
    Dim dPadDepth As Double
    Dim dPadWidth As Double
    Dim dTapBodyWidth As Double
    Dim dPadBodyDepth As Double
    
    If parHVACShape = FlatOval Then
        Dim oFltOvalCurve As Object
        'Create the non persistent Duct (Header)
        oCenter.Set -(dBWidth / 2 + 10), 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOval(oCenter, parWidth, _
                                            parDepth, 0)
        Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    1, 0, 0, _
                                                    (dBWidth + 20), False)
        Set oFltOvalCurve = Nothing

        If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
            If CmpDblLessThan(dBDepth + 2 * INCH, parDepth) Then
                'Create the Non-Persistant TapBody
                oCenter.Set 0, 0, -parDepth / 2
                
                Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, dBWidth + 2 * parInsulationThickness, _
                                                            dBDepth + 2 * parInsulationThickness, _
                                                                        0, 0, parAngle)
                oDirVector.Set Cos(parAngle), -Sin(parAngle), 0
                oDirVector.Length = 1
                Set oTapBody = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    oDirVector.X, oDirVector.Y, oDirVector.Z, _
                                                    dTapLength, False)
                Set oFltOvalCurve = Nothing
                
                oDirVector.Set Cos(parAngle), -Sin(parAngle), 0
                oDirVector.Length = 1
                oCenter.Set 0, 0, -parDepth / 2
                
                Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, oCenter.X, _
                                                 oCenter.Y, oCenter.Z, _
                                                oDirVector.X, oDirVector.Y, oDirVector.Z)
                oCenter.Set dTapLength * Cos(parAngle), _
                                    -dTapLength * Sin(parAngle), -parDepth / 2
                
                oTransformationMat.LoadIdentity
                oDirVector.Set oCenter.X, oCenter.Y, oCenter.Z
                oTransformationMat.Translate oDirVector
                oPlane.Transform oTransformationMat
            Else
                'Create the Non-Persistant TapBody
                oCenter.Set 0, 0, -parDepth / 2
                
                Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, dBWidth + 2 * parInsulationThickness, _
                                                            dBDepth + 2 * parInsulationThickness, _
                                                                        0, 0, parAngle)
                oDirVector.Set Cos(parAngle), -Sin(parAngle), 0
                oDirVector.Length = 1
                Set oTapBody = oGeomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                    oFltOvalCurve, _
                                                    oDirVector.X, oDirVector.Y, oDirVector.Z, _
                                                    dTapLength, False)
                Set oFltOvalCurve = Nothing
            End If

        ElseIf lPartdatabasis = 80 Then
            
            Dim oFltOvalCurve1 As Object
            Dim oFltOvalCurve2 As Object
            Dim dTapBodyDepth As Double

            dTapBodyDepth = dBDepth + 1 * INCH
            If CmpDblLessThan(dBDepth + 2 * parInsulationThickness, parDepth) Then
                'Create the non persistent Duct (Header) by decreasing the
                'Y to Y-parInsulationThickness to get the correct intersection curve
                oCenter.Set -(dBWidth / 2 + 10), parInsulationThickness, -parDepth / 2
                Set oFltOvalCurve = CreFlatOval(oCenter, parWidth + 2 * parInsulationThickness, _
                                                    parDepth + 2 * parInsulationThickness, 0)
                Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                            1, 0, 0, _
                                                            (dBWidth + 20), False)
                        
                'Create the Non-persistant TapBody
                oCenter.Set 0, 0, -parDepth / 2
                Set oFltOvalCurve1 = FltOvlBranchTransCurve(oCenter, _
                                    dBWidth + 1 * INCH + 2 * parInsulationThickness, _
                                    dTapBodyDepth + 2 * parInsulationThickness, 0, 0, parAngle)
                oCenter.Set 0, -(dTapLength), -parDepth / 2
                Set oFltOvalCurve2 = FltOvlBranchTransCurve(oCenter, _
                                        dBWidth + 2 * parInsulationThickness, _
                                        parBDepth + 2 * parInsulationThickness, 0, 0, parAngle)
    
                Set oTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oFltOvalCurve1, _
                                                oFltOvalCurve2, True)
    
                'Create a Plane normal to the Branch to get the flat oval curve at the nozzle
                
                oCenter.Set 0, 0, -parDepth / 2
                oDirVector.Set Cos(parAngle), -Sin(parAngle), 0
                oDirVector.Length = 1
                
                Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, oCenter.X, oCenter.Y, _
                                            oCenter.Z, oDirVector.X, oDirVector.Y, oDirVector.Z)
                oCenter.Set dTapLength * Cos(parAngle), _
                                    -dTapLength * Sin(parAngle), -parDepth / 2
    
                oTransformationMat.LoadIdentity
                oDirVector.Set oCenter.X, oCenter.Y, oCenter.Z
                oTransformationMat.Translate oDirVector
                oPlane.Transform oTransformationMat
            Else
                oCenter.Set 0, 0, -parDepth / 2
                Set oFltOvalCurve1 = FltOvlBranchTransCurve(oCenter, _
                                    dBWidth + 1 * INCH + 2 * parInsulationThickness, _
                                    dTapBodyDepth + 2 * parInsulationThickness, 0, 0, parAngle)
                oCenter.Set 0, -(dTapLength), -parDepth / 2
                Set oFltOvalCurve2 = FltOvlBranchTransCurve(oCenter, _
                                        dBWidth + 2 * parInsulationThickness, _
                                        parBDepth + 2 * parInsulationThickness, 0, 0, parAngle)
    
                Set oTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oFltOvalCurve1, oFltOvalCurve2, True)
            End If
            
            Set oFltOvalCurve1 = Nothing
            Set oFltOvalCurve2 = Nothing
        End If

        If CmpDblLessThan(dBDepth + 2 * INCH, parDepth) Then
            'Curve between TapBody and Header
            Set oIJSurfaceHeader = oTapBody
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve2 = oIntersectionCurve.Item(1)
            
            'Curve at the of the nozzle
            Set oIJSurfaceHeader = oPlane
            oIJSurfaceHeader.Intersect oTapBody, oIntersectionCurve, IntersectCode
            Set oIJcurve3 = oIntersectionCurve.Item(1)

            'Create Tap Body
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oIJcurve2, oIJcurve3, True)
        Else
            Set objTapBody = oTapBody
        End If
            

        'Create the Take Off
        oCenter.Set 0, 0, -parDepth / 2
        Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, _
                                1.01 * dBWidth + 2 * parInsulationThickness, _
                                dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness, _
                                0, 0, parAngle)
        oDirVector.Set Cos(parAngle), -Sin(parAngle), 0
        Set objTakeOff = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 2 * INCH, True)
        oCenter.Set dTapLength * Cos(parAngle), _
                                -dTapLength * Sin(parAngle), 0
        oTransformationMat.LoadIdentity
        oDirVector.Set oCenter.X, oCenter.Y, oCenter.Z
        oTransformationMat.Translate oDirVector
        objTakeOff.Transform oTransformationMat
        Set oFltOvalCurve = Nothing

        'Create Strip
        oCenter.Set 0, 0, 0
        Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, _
                                    1.06 * dBWidth + 2 * parInsulationThickness, _
                                    dBDepth + 0.06 * dBDepth + 2 * parInsulationThickness, _
                                    0, 0, parAngle)
        Set objStrip = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 0.008, True)
        oCenter.Set (dTapLength - 0.004) * Cos(parAngle), _
                                -(dTapLength - 0.004) * Sin(parAngle), -parDepth / 2
        oTransformationMat.LoadIdentity
        oDirVector.Set oCenter.X, oCenter.Y, oCenter.Z
        oTransformationMat.Translate oDirVector
        objStrip.Transform oTransformationMat
        Set oFltOvalCurve = Nothing
        
        'But in this case translate all objects to a distance of (parWidth/2)/tan(parAngle)
        'in -ve X direction to get the correct offset from origin
        oTransformationMat.LoadIdentity
        oDirVector.Set -parWidth / (2 * Tan(parAngle)), 0, 0
        oTransformationMat.Translate oDirVector
        objTapBody.Transform oTransformationMat
        objStrip.Transform oTransformationMat
        objTakeOff.Transform oTransformationMat
    
    ElseIf parHVACShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
            
            'Create the Tap Body
            oCenter.Set 0, 0, 0
            Set oBotCurve = CreSMRectBranch(oCenter, _
                                    ((dBWidth + 2 * parInsulationThickness) / Sin(parAngle)), _
                                    dBDepth + 2 * parInsulationThickness, 0, 0)
            
            oCenter.Set 0, 0, dTapLenFrmOrg
            Set oTopCurve = CreSMRectBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                            dBDepth + 2 * parInsulationThickness, 0, parAngle)
            
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oTopCurve, oBotCurve, True)

        ElseIf lPartdatabasis = 80 Then

            'Create the Tap Body
            oCenter.Set 0, 0, 0
            Set oBotCurve = CreSMRectBranch(oCenter, _
                                ((dBWidth + 1 * INCH + 2 * parInsulationThickness) / Sin(parAngle)), _
                                dBDepth + 1 * INCH + 2 * parInsulationThickness, 0, 0)
            
            oCenter.Set 0, 0, dTapLenFrmOrg
            Set oTopCurve = CreSMRectBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                        dBDepth + 2 * parInsulationThickness, 0, parAngle)
            
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oTopCurve, oBotCurve, True)
        End If

        'Create Take Off
        oStPoint.Set (dTapLenFrmOrg) * Cos(parAngle), 0, dTapLenFrmOrg * Sin(parAngle)
        Set objTakeOff = PlaceTrapezoid(m_OutputColl, oStPoint, _
                                    1.01 * dBWidth + 2 * parInsulationThickness, _
                                    dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness, _
                                    1.01 * dBWidth + 2 * parInsulationThickness, _
                                    dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness, _
                                    2 * INCH, True, 0, parAngle, 0)

        'Create Strip
        
        oStPoint.Set (dTapLenFrmOrg + 2 * INCH - 0.004) * Cos(parAngle), 0, (dTapLenFrmOrg + 2 * INCH - 0.004) * Sin(parAngle)
        Set objStrip = PlaceTrapezoid(m_OutputColl, oStPoint, _
                                    1.06 * dBWidth + 2 * parInsulationThickness, _
                                    dBDepth + 0.06 * dBWidth + 2 * parInsulationThickness, _
                                    1.06 * dBWidth + 2 * parInsulationThickness, _
                                    dBDepth + 0.06 * dBWidth + 2 * parInsulationThickness, _
                                    0.004, True, 0, parAngle, 0)
        
    ElseIf parHVACShape = 4 Then 'Round
        'Create non-prsistant Header having length greater than the TapBody Width
        
        oStPoint.Set -(dBWidth + 10), 0, -parWidth / 2
        oEnPoint.Set (dBWidth + 10), 0, -parWidth / 2
        Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth)
        
        If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
            If CmpDblLessThan(dBWidth + 2 * parInsulationThickness, parWidth) Then
                'Create the Tap Body
                oStPoint.Set -(parWidth / 2) / Tan(parAngle), 0, -(parWidth / 2)
                oEnPoint.Set dTapLenFrmOrg * Cos(parAngle), 0, (dTapLenFrmOrg * Sin(parAngle))
                
                Set oTapBody = PlaceCylinderTrans(oStPoint, oEnPoint, _
                                                    (dBWidth + 2 * parInsulationThickness))
                oCenter.Set dTapLenFrmOrg * Cos(parAngle), 0, (dTapLenFrmOrg * Sin(parAngle))
                oDirVector.Set -Cos(parAngle), 0, -Sin(parAngle)
                oDirVector.Length = 1
                Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, _
                                                     (dBWidth + 2 * parInsulationThickness) / 2)
    
                oTransformationMat.LoadIdentity
                dRotAboutZaxis = PI / 2
                oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
                oTransformationMat.Rotate dRotAboutZaxis, oDirVector
                oTopCurve.Transform oTransformationMat
            Else
                oStPoint.Set -(parWidth / 2) / Tan(parAngle), 0, -(parWidth / 2)
                oEnPoint.Set dTapLenFrmOrg * Cos(parAngle), 0, (dTapLenFrmOrg * Sin(parAngle))
                Set oTapBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBWidth + 2 * parInsulationThickness, True)
            End If
        ElseIf lPartdatabasis = 80 Then
            
            Dim dTapBotWidth As Double
            'Create the non-persistant Tap Body
            'Create the Cirlce at the middle of the Duct
            dTapBotWidth = dBWidth + ((dTapLength * dAlpha) / (2 * INCH))
            
            If CmpDblLessThan(dBWidth + 2 * parInsulationThickness, parWidth) Then
                'Create the Header by decreasing the -Z to -(Z + parInsulationThickness) _
                to get intersection curve correctly
                oStPoint.Set -(dBWidth + 10), 0, -(parWidth / 2 + parInsulationThickness)
                oEnPoint.Set (dBWidth + 10), 0, -(parWidth / 2 + parInsulationThickness)
                Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth + 2 * parInsulationThickness)
                
                If CmpDblGreaterthan(dTapBotWidth, parWidth) Then
                    dTapBotWidth = parWidth - 0.001
                End If
                'Create the top and bottom Circles for tap body
                'Circle at the middle of the duct
                oCenter.Set -(parWidth / (2 * Tan(parAngle))), 0, -parWidth / 2
                oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
                oDirVector.Length = 1
                Set oBotCurve = PlaceTrCircleByCenter(oCenter, oDirVector, _
                                                (dTapBotWidth + 2 * parInsulationThickness) / 2)
                
                'Create the Circle at the Nozzle
                oCenter.Set dTapLenFrmOrg * Cos(parAngle), 0, dTapLenFrmOrg * Sin(parAngle)
                oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
                oDirVector.Length = 1
                Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, _
                                                    (dBWidth + 2 * parInsulationThickness) / 2)
                
                'Create the non-persistant Tap Body
                Set oTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oBotCurve, oTopCurve, True)
                
                'Create a non-persistant Plane at the Nozzle to get the intersection curve
                'betwwen Tap Body and this Plane
                Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                        oCenter.X, oCenter.Y, oCenter.Z, _
                                                        oDirVector.X, oDirVector.Y, oDirVector.Z)
                Set oIJSurfaceHeader = oTapBody
                oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
                Set oIJcurve3 = oIntersectionCurve.Item(1)
            
            Else
                oStPoint.Set -(parWidth / (2 * Tan(parAngle))), 0, -parWidth / 2
                oEnPoint.Set dTapLenFrmOrg * Cos(parAngle), 0, dTapLenFrmOrg * Sin(parAngle)
                Set oTapBody = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                            (dTapBotWidth + 2 * parInsulationThickness) / 2, _
                                            (dBWidth + 2 * parInsulationThickness) / 2, True)
            End If
           
        End If
        If CmpDblLessThan(dBWidth + 2 * parInsulationThickness, parWidth) Then
            Set oIJSurfaceHeader = oTapBody
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve2 = oIntersectionCurve.Item(1)
         End If
        'Create Tap Body
        If lPartdatabasis <= 1 Or lPartdatabasis = 75 Then
            If CmpDblLessThan(dBWidth + 2 * parInsulationThickness, parWidth) Then
                Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oTopCurve, oIJcurve2, True)
            Else
                Set objTapBody = oTapBody
            End If
        ElseIf lPartdatabasis = 80 Then
            If CmpDblLessThan(dBWidth + 2 * parInsulationThickness, parWidth) Then
                Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oIJcurve3, oIJcurve2, True)
            Else
                Set objTapBody = oTapBody
            End If
        End If
        
        'Create Take Off
        oStPoint.Set (dTapLenFrmOrg) * Cos(parAngle), 0, _
                        ((dTapLenFrmOrg) * Sin(parAngle))
        oEnPoint.Set (dTapLenFrmOrg + 2 * INCH) * Cos(parAngle), 0, _
                        ((dTapLenFrmOrg + 2 * INCH) * Sin(parAngle))
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        1.01 * dBWidth + 2 * parInsulationThickness, True)
        
        'Create Strip
        oStPoint.Set (dTapLenFrmOrg + 2 * INCH - 0.004) * Cos(parAngle), 0, _
                        ((dTapLenFrmOrg + 2 * INCH - 0.004) * Sin(parAngle))
        oEnPoint.Set (dTapLenFrmOrg + 2 * INCH + 0.004) * Cos(parAngle), 0, _
                        ((dTapLenFrmOrg + 2 * INCH + 0.004) * Sin(parAngle))
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                1.06 * dBWidth + 2 * parInsulationThickness, True)
    
    End If
    
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
    Set oIJcurve2 = Nothing
    Set oIJcurve3 = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext

End Sub





