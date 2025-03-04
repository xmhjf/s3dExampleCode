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
'   ProgID:         SP3DHContFlgdLateralTap.HCFLTap
'   Author:         RUK
'   Creation Date:  Wednesday, June 27 2007
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
    Dim dNozzWidth As Double
    Dim dNozzDepth As Double
    
    dBWidth = parBWidth
    dBDepth = parBDepth
    dNozzWidth = dBWidth
    dNozzDepth = dBDepth
    
    If parHVACShape = 4 Then 'Round
        If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
            If CmpDblGreaterthan(parBWidth, parWidth) Then
                dBWidth = parWidth - 0.002
                dNozzWidth = parWidth
            End If
        ElseIf lPartdatabasis = 70 Then
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
            dNozzWidth = dBWidth
        End If
    ElseIf parHVACShape = FlatOval Or parHVACShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
            If CmpDblGreaterthanOrEqualTo(parBDepth, parDepth) Then
                dBDepth = parDepth - 0.002
                dNozzDepth = parDepth
            End If
            If CmpDblGreaterthanOrEqualTo(parBWidth, parWidth) Then
                dBWidth = parWidth - 0.002
                dNozzWidth = parWidth
            End If
        
        ElseIf lPartdatabasis = 70 Then
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
            dNozzWidth = dBWidth
            dNozzDepth = dBDepth
        End If
    End If
    
    'Caliculate the Length of the Tap as per McGill HVAC Catalog
    If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
        dTapLength = (parWidth / (2 * Sin(parAngle))) + _
                            (dBWidth / (2 * Tan(parAngle))) + 2 * INCH
    ElseIf lPartdatabasis = 70 Then
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
    Dim oBotCurve As Object
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
        oCenter.Set -((dBWidth / 2) + 10), 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOval(oCenter, parWidth, parDepth, 0)
        Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    1, 0, 0, _
                                                    (dBWidth + 20), False)
        Set oFltOvalCurve = Nothing

        If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
            'Create the Non-persistant Pad
            oCenter.Set 0, 0, -parDepth / 2
            dPadDepth = dBDepth + 1 * INCH
            If CmpDblGreaterthan(dPadDepth, parDepth) Then
                dPadDepth = parDepth - 0.001
            End If
            dPadWidth = dBWidth + 1 * INCH
            If CmpDblGreaterthan(dPadWidth, parWidth) Then
                dPadWidth = parWidth - 0.001
            End If
            Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, dPadWidth, dPadDepth, _
                                            0, 0, parAngle)
            
            oDirVector.Set Cos(parAngle), -Sin(parAngle), 0
            oDirVector.Length = 1
            Set oPad = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    oDirVector.X, oDirVector.Y, oDirVector.Z, _
                                                    dTapLength, False)
            Set oFltOvalCurve = Nothing

            'Create the Non-Persistant TapBody
            oCenter.Set 0, 0, -parDepth / 2
            
            Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, dBWidth, dBDepth, _
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

        ElseIf lPartdatabasis = 70 Then
            Dim oFltOvalCurve1 As Object
            Dim oFltOvalCurve2 As Object
            Dim dTapBodyDepth As Double

            'Create the Non-persistant Pad
            oCenter.Set 0, 0, -parDepth / 2
            dPadDepth = dBDepth + 2 * INCH
            If CmpDblGreaterthan(dPadDepth, parDepth) Then
                dPadDepth = parDepth - 0.001
            End If
            dPadWidth = dBWidth + 2 * INCH
            If CmpDblGreaterthanOrEqualTo(dPadWidth, parWidth) Then
                dPadWidth = parWidth - 0.001
            End If
            Set oFltOvalCurve1 = FltOvlBranchTransCurve(oCenter, dPadWidth, dPadDepth, _
                                                        0, 0, parAngle)

            oCenter.Set 0, -(dTapLength), -parDepth / 2
            Set oFltOvalCurve2 = FltOvlBranchTransCurve(oCenter, dPadWidth - 1 * INCH, _
                                                dPadDepth - 1 * INCH, 0, 0, parAngle)
            
            Set oPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oFltOvalCurve1, _
                                                    oFltOvalCurve2, True)

            'Create the Non-persistant TapBody
            oCenter.Set 0, 0, -parDepth / 2
            dTapBodyDepth = dBDepth + 1 * INCH
            If CmpDblGreaterthan(dTapBodyDepth, parDepth) Then
                dTapBodyDepth = parDepth - 0.002
            End If
            dTapBodyWidth = dBWidth + 1 * INCH
            If CmpDblGreaterthanOrEqualTo(dTapBodyWidth, parWidth) Then
                dTapBodyWidth = parWidth - 0.002
            End If
            Set oFltOvalCurve1 = FltOvlBranchTransCurve(oCenter, dTapBodyWidth, _
                                                            dTapBodyDepth, 0, 0, parAngle)
            oCenter.Set 0, -(dTapLength), -parDepth / 2
            Set oFltOvalCurve2 = FltOvlBranchTransCurve(oCenter, dBWidth, parBDepth, 0, 0, parAngle)

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
            
            Set oFltOvalCurve1 = Nothing
            Set oFltOvalCurve2 = Nothing
        End If


'        Curve between Pad and Header
        Set oIJSurfaceHeader = oPad
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve1 = oIntersectionCurve.Item(1)

        'Curve between TapBody and Header
        Set oIJSurfaceHeader = oTapBody
        oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
        Set oIJcurve2 = oIntersectionCurve.Item(1)

'        Curve at the of the nozzle
        Set oIJSurfaceHeader = oPlane
        oIJSurfaceHeader.Intersect oTapBody, oIntersectionCurve, IntersectCode
        Set oIJcurve3 = oIntersectionCurve.Item(1)

        'Create Pad
        Set objPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oIJcurve1, oIJcurve2, True)
        'Create Tap Body
        Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oIJcurve2, oIJcurve3, True)
        
        'Create the Take Off
        oCenter.Set 0, 0, -parDepth / 2
        Set oFltOvalCurve = FltOvlBranchTransCurve(oCenter, _
                                1.01 * dBWidth, dBDepth + 0.01 * dBWidth, 0, 0, parAngle)
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
                                    1.06 * dBWidth, dBDepth + 0.06 * dBDepth, 0, 0, parAngle)
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
        objPad.Transform oTransformationMat
        objStrip.Transform oTransformationMat
        objTakeOff.Transform oTransformationMat
    
    ElseIf parHVACShape = Rectangular Then
        If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
            
            'Create the Pad
            oCenter.Set 0, 0, 0
            Set oBotCurve = CreSMRectBranch(oCenter, (dBWidth / Sin(parAngle)) + 1 * INCH, dBDepth + 1 * INCH, 0, 0)
            
            oDirVector.Set 0, 0, 1
            Set objPad = PlaceProjection(m_OutputColl, oBotCurve, oDirVector, 0.001, True)
            
            'Create the Tap Body
            oCenter.Set 0, 0, 0
            Set oBotCurve = CreSMRectBranch(oCenter, dBWidth / Sin(parAngle), dBDepth, 0, 0)
            
            oCenter.Set 0, 0, dTapLenFrmOrg
            Set oTopCurve = CreSMRectBranch(oCenter, dBWidth, dBDepth, 0, parAngle)
            
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oTopCurve, oBotCurve, True)

        ElseIf lPartdatabasis = 70 Then

            'Create the Pad
            oCenter.Set 0, 0, 0
            Set oBotCurve = CreSMRectBranch(oCenter, (dBWidth / Sin(parAngle)) + 2 * INCH, dBDepth + 2 * INCH, 0, 0)
            
            oDirVector.Set 0, 0, 1
            Set objPad = PlaceProjection(m_OutputColl, oBotCurve, oDirVector, 0.001, True)

            'Create the Tap Body
            oCenter.Set 0, 0, 0
            Set oBotCurve = CreSMRectBranch(oCenter, dBWidth / Sin(parAngle) + 1 * INCH, dBDepth + 1 * INCH, 0, 0)
            
            oCenter.Set 0, 0, dTapLenFrmOrg
            Set oTopCurve = CreSMRectBranch(oCenter, dBWidth, dBDepth, 0, parAngle)
            
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oTopCurve, oBotCurve, True)
        End If

        'Create Take Off
        oStPoint.Set (dTapLenFrmOrg) * Cos(parAngle), 0, dTapLenFrmOrg * Sin(parAngle)
        Set objTakeOff = PlaceTrapezoid(m_OutputColl, oStPoint, 1.01 * dBWidth, dBDepth + 0.01 * dBWidth, _
                                        1.01 * dBWidth, dBDepth + 0.01 * dBWidth, 2 * INCH, True, 0, parAngle, 0)

        'Create Strip
        
        oStPoint.Set (dTapLenFrmOrg + 2 * INCH - 0.004) * Cos(parAngle), 0, (dTapLenFrmOrg + 2 * INCH - 0.004) * Sin(parAngle)
        Set objStrip = PlaceTrapezoid(m_OutputColl, oStPoint, 1.06 * dBWidth, dBDepth + 0.06 * dBWidth, _
                                         1.06 * dBWidth, dBDepth + 0.06 * dBWidth, 0.004, True, 0, parAngle, 0)
        
    ElseIf parHVACShape = 4 Then 'Round
        'Create non-prsistant Header having length greater than the TapBody Width
        
        oStPoint.Set -(dBWidth + 10), 0, -parWidth / 2
        oEnPoint.Set (dBWidth + 10), 0, -parWidth / 2
        Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth)
        
        If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
            'Create the Tap Body
            oStPoint.Set -(parWidth / 2) / Tan(parAngle), 0, -(parWidth / 2)
            oEnPoint.Set dTapLenFrmOrg * Cos(parAngle), 0, (dTapLenFrmOrg * Sin(parAngle))
            
            Set oTapBody = PlaceCylinderTrans(oStPoint, oEnPoint, dBWidth)
            
            dPadWidth = dBWidth + 1 * INCH
            If CmpDblGreaterthanOrEqualTo(dPadWidth, parWidth) Then
                dPadWidth = parWidth - 0.001
            End If
            Set oPad = PlaceCylinderTrans(oStPoint, oEnPoint, dPadWidth)
            
            oCenter.Set dTapLenFrmOrg * Cos(parAngle), 0, (dTapLenFrmOrg * Sin(parAngle))
            oDirVector.Set -Cos(parAngle), 0, -Sin(parAngle)
            oDirVector.Length = 1
            Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, dBWidth / 2)

            oTransformationMat.LoadIdentity
            dRotAboutZaxis = PI / 2
            oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
            oTransformationMat.Rotate dRotAboutZaxis, oDirVector
            oTopCurve.Transform oTransformationMat
            
        ElseIf lPartdatabasis = 70 Then
            
            'Create the Non-persistant Pad
            'Create the Bottom Circle at the middle Of the duct
            Dim dTapBotWidth As Double
            dTapBotWidth = dBWidth + ((dTapLength * dAlpha) / (2 * INCH)) + 1 * INCH
            
            If CmpDblGreaterthan(dTapBotWidth, parWidth) Then
                dTapBotWidth = (parWidth) - 0.002
            End If
            oCenter.Set -(parWidth / (2 * Tan(parAngle))), 0, -parWidth / 2
            oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
            oDirVector.Length = 1
            Set oBotCurve = PlaceTrCircleByCenter(oCenter, oDirVector, dTapBotWidth / 2)

            'Create the Circle at the Nozzle
            oCenter.Set dTapLenFrmOrg * Cos(parAngle), 0, dTapLenFrmOrg * Sin(parAngle)
            oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
            oDirVector.Length = 1
            Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, (dBWidth + 1 * INCH) / 2)
           
            'Create the non-persistant Tap Body
            Set oPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oBotCurve, oTopCurve, True)
            
            'Create the non-persistant Tap Body
            'Create the Cirlce at the middle of the Duct
            dTapBotWidth = dBWidth + ((dTapLength * dAlpha) / (2 * INCH))
            oCenter.Set -(parWidth / (2 * Tan(parAngle))), 0, -parWidth / 2
            oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
            oDirVector.Length = 1
            If CmpDblGreaterthanOrEqualTo(dTapBotWidth, parWidth) Then
                dTapBotWidth = parWidth - 0.001
            End If
            Set oBotCurve = PlaceTrCircleByCenter(oCenter, oDirVector, dTapBotWidth / 2)
            
            'Create the Circle at the Nozzle
            oCenter.Set dTapLenFrmOrg * Cos(parAngle), 0, dTapLenFrmOrg * Sin(parAngle)
            oDirVector.Set Cos(parAngle), 0, Sin(parAngle)
            oDirVector.Length = 1
            Set oTopCurve = PlaceTrCircleByCenter(oCenter, oDirVector, dBWidth / 2)
            
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
           
        End If

            Set oIJSurfaceHeader = oPad
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve1 = oIntersectionCurve.Item(1)
            
            Set oIJSurfaceHeader = oTapBody
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve2 = oIntersectionCurve.Item(1)
            
            'Create the Pad
            Set objPad = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oIJcurve1, oIJcurve2, False)
            
            'Create Tap Body
            If lPartdatabasis <= 1 Or lPartdatabasis = 65 Then
                Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oTopCurve, oIJcurve2, True)
            ElseIf lPartdatabasis = 70 Then
                Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oIJcurve3, oIJcurve2, True)
            End If
            
            'Create Take Off
            oStPoint.Set (dTapLenFrmOrg) * Cos(parAngle), 0, _
                            ((dTapLenFrmOrg) * Sin(parAngle))
            oEnPoint.Set (dTapLenFrmOrg + 2 * INCH) * Cos(parAngle), 0, _
                            ((dTapLenFrmOrg + 2 * INCH) * Sin(parAngle))
            Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * dBWidth, True)
            
            'Create Strip
            oStPoint.Set (dTapLenFrmOrg + 2 * INCH - 0.004) * Cos(parAngle), 0, _
                            ((dTapLenFrmOrg + 2 * INCH - 0.004) * Sin(parAngle))
            oEnPoint.Set (dTapLenFrmOrg + 2 * INCH + 0.004) * Cos(parAngle), 0, _
                            ((dTapLenFrmOrg + 2 * INCH + 0.004) * Sin(parAngle))
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
        dDepth = dNozzDepth
        dWidth = dNozzWidth
        lCSType = GSCADNozzleEntities.Rectangular
    ElseIf parHVACShape = 4 Then 'Round
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
        dDepth = dNozzWidth
        dWidth = dNozzWidth
        lCSType = GSCADNozzleEntities.Round
    ElseIf parHVACShape = FlatOval Then
        dDepth = dNozzDepth
        dWidth = dNozzWidth
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
        X = dTapLenFrmOrg * Cos(parAngle)
        Y = 0
        Z = (dTapLenFrmOrg * Sin(parAngle))
    ElseIf parHVACShape = FlatOval Then
        X = dTapLength * Cos(parAngle) - parWidth / (2 * Tan(parAngle))
        Y = -(dTapLength * Sin(parAngle))
        Z = -parDepth / 2
    ElseIf parHVACShape = Rectangular Then
        X = dTapLenFrmOrg * Cos(parAngle)
        Y = 0
        Z = dTapLenFrmOrg * Sin(parAngle)
    End If
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set Cos(parAngle), 0, Sin(parAngle)
    dir.Length = 1
    oDistribPort.SetDirectionVector dir
    If parHVACShape = FlatOval Then
        dir.Set Cos(parAngle), -Sin(parAngle), 0
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
End Sub

