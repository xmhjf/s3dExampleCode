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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'       This symbol is taken from Carnes HVAC Catalog.
'
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------     ---      ------------------
'   29.Jun.2006     svsmylav TR-100615    Corrected baffle gap calculation logic to avoid negative distance values.
'                                         Also removed 'Depth = Width' statement as it is not required for rectangular port.
'  08.SEP.2006      KKC      DI-95670     Replace names with initials in all revision history sheets and symbols
'  22.SEP.2008      PK       CR-149529    Modified the symbol code for the partdatabasis cases 161,162,163,164
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private RAD          As Double
Const NEGLIGIBLE_GAP = 0.0001
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput                 As Double
    Dim ObjDiffuserNeck         As Object
    Dim ObjDiffuserProjection   As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle2      As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle3      As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle4      As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle5      As IngrGeom3D.RuledSurface3d
    
    Dim Width                   As Double
    Dim Depth                   As Double
    
    'Here BWidth & BDepth are the Diffuser's Width and Depth  at the larger side.
    'No port is being routed from this end.

    Dim BWidth                  As Double
    Dim BDepth                  As Double
    
    Dim DiffuserThickness       As Double
    Dim BaffleGap               As Double
    Dim oPos                    As New AutoMath.DPosition
    Dim oDir                    As New AutoMath.DVector
    
    Dim dblPnts1(0 To 11)       As Double
    Dim dblPnts2(0 To 11)       As Double
    
    Dim DiffuserLength1         As Double
        
    Dim geomFactory    As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'Width = arrayOfInputs(2) 'A
    'Depth = arrayOfInputs(3) 'B
    'BWidth = arrayOfInputs(4)
    'BDepth = arrayOfInputs(5)
    'DiffuserThickness = arrayOfInputs(6)
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    'Partdatabasis cases implemented
    'Rectangular diffuser, with rectangular neck, Type A-161
    'Rectangular diffuser, with rectangular neck, Type B -162
    'Rectangular diffuser, with rectangular neck, Type C -163
    'Rectangular diffuser, with rectangular neck, Type D -164
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 161       'Rectangular diffuser, with rectangular neck, Type A
            Width = arrayOfInputs(2) 'A
            Depth = arrayOfInputs(3) 'B
            BWidth = arrayOfInputs(4)
            BDepth = arrayOfInputs(5)
            DiffuserThickness = arrayOfInputs(6)
        
        Case 162                'Rectangular diffuser, with rectangular neck, Type B
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            BWidth = arrayOfInputs(4)
            BDepth = arrayOfInputs(5)
            DiffuserThickness = arrayOfInputs(6)
            
        Case 163                'Rectangular diffuser, with rectangular neck, Type C
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            BWidth = arrayOfInputs(4)
            BDepth = arrayOfInputs(5)
            DiffuserThickness = arrayOfInputs(6)
            
        Case 164                'Rectangular diffuser, with rectangular neck, Type D
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            BWidth = arrayOfInputs(4)
            BDepth = arrayOfInputs(5)
            DiffuserThickness = arrayOfInputs(6)
            DiffuserLength1 = arrayOfInputs(7)
                                    
    End Select
    
    Select Case lPartDataBasis
     
        Case Is <= 1

            BaffleGap = (25.4 / 1000) * 1.5
            
            Dim Width5 As Double
            Dim Depth5 As Double
            Dim StripThickness As Double
            
            StripThickness = 10 / 1000  ' Strip Thickness standard value = 0.15625 inches
            
            Width5 = BWidth - 2 * StripThickness
            Depth5 = BDepth - 2 * StripThickness
            
               
            Dim Width4 As Double
            Dim Depth4 As Double
            
            Dim Width3 As Double
            Dim Depth3 As Double
            
        
            Dim Width1 As Double
            Dim Depth1 As Double
            
            Dim Width0 As Double
            Dim Depth0 As Double
            
            Dim Width2 As Double
            Dim Depth2 As Double
            
            Dim Width21 As Double
            Dim Depth21 As Double
        
            Width4 = Width5 - BaffleGap * 2
            Depth4 = Depth5 - BaffleGap * 2
            
            Width3 = Width4 - BaffleGap * 2
            Depth3 = Depth4 - BaffleGap * 2
            
            Width2 = Width3 - BaffleGap * 2
            Depth2 = Depth3 - BaffleGap * 2
               
            'Compute width/depth values at inlet
            Dim dInletWidthBaffleGap As Double
            Dim dInletDepthBaffleGap As Double
            
            dInletWidthBaffleGap = Width / 6
            If CmpDblGreaterthan(dInletWidthBaffleGap, BaffleGap) Then dInletWidthBaffleGap = BaffleGap
            dInletDepthBaffleGap = Depth / 6
            If CmpDblGreaterthan(dInletDepthBaffleGap, BaffleGap) Then dInletDepthBaffleGap = BaffleGap
        
            Width0 = Width - dInletWidthBaffleGap * 2
            Depth0 = Depth - dInletDepthBaffleGap * 2
        
            Width1 = Width0 - dInletWidthBaffleGap * 2
            Depth1 = Depth0 - dInletDepthBaffleGap * 2
        
            Width21 = Width1 - dInletWidthBaffleGap * 2
            Depth21 = Depth1 - dInletDepthBaffleGap * 2
        
            If CmpDblLessThanOrEqualTo(Width21, NEGLIGIBLE_GAP) Then
                Width21 = NEGLIGIBLE_GAP
            End If
            If CmpDblLessThanOrEqualTo(Depth21, NEGLIGIBLE_GAP) Then
                Depth21 = NEGLIGIBLE_GAP
            End If
            
            iOutput = 0
        
            oPos.Set 0, 0, 0
            
            'Insert your code for output 1(Diffuser Neck)
          
            Dim stPoint   As New AutoMath.DPosition
            Dim enPoint   As New AutoMath.DPosition
            
            stPoint.Set 0, -Depth / 2, -Width / 2
            enPoint.Set DiffuserThickness / 2, Depth / 2, Width / 2
            
            Set ObjDiffuserNeck = PlaceBox(m_OutputColl, stPoint, enPoint)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserNeck
            Set ObjDiffuserNeck = Nothing
            
            'Insert your code for output 2(Neck Projection)
            
            stPoint.Set DiffuserThickness / 2, 0, 0
        
            Set ObjDiffuserProjection = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width, Depth, Width5, Depth5, _
                            DiffuserThickness / 2, False, 0, PI / 2, 0)
                            
            'Set the output
            m_OutputColl.AddOutput "NeckProjection", ObjDiffuserProjection
            Set ObjDiffuserProjection = Nothing
            
            
            'Insert your code for output 3(Diffuser Stiffner -1)
            Dim ObjOuterPlane As IngrGeom3D.Plane3d
             
            oPos.Set DiffuserThickness, -BDepth / 2, -BWidth / 2
            
            dblPnts1(0) = oPos.X
            dblPnts1(1) = oPos.Y
            dblPnts1(2) = oPos.Z
            
            dblPnts1(3) = oPos.X
            dblPnts1(4) = oPos.Y + BDepth
            dblPnts1(5) = oPos.Z
            
            dblPnts1(6) = oPos.X
            dblPnts1(7) = oPos.Y + BDepth
            dblPnts1(8) = oPos.Z + BWidth
            
            dblPnts1(9) = oPos.X
            dblPnts1(10) = oPos.Y
            dblPnts1(11) = oPos.Z + BWidth
            
            Set ObjOuterPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblPnts1)
            
            oPos.Set DiffuserThickness, -Depth5 / 2, -Width5 / 2
            
            dblPnts2(0) = oPos.X
            dblPnts2(1) = oPos.Y
            dblPnts2(2) = oPos.Z
            
            dblPnts2(3) = oPos.X
            dblPnts2(4) = oPos.Y
            dblPnts2(5) = oPos.Z + Width5
            
            dblPnts2(6) = oPos.X
            dblPnts2(7) = oPos.Y + Depth5
            dblPnts2(8) = oPos.Z + Width5
            
            dblPnts2(9) = oPos.X
            dblPnts2(10) = oPos.Y + Depth5
            dblPnts2(11) = oPos.Z
        
            Dim oLine1           As IngrGeom3D.Line3d
            Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(0), dblPnts2(1), dblPnts2(2), dblPnts2(3), dblPnts2(4), dblPnts2(5))
            
            Dim oLine2           As IngrGeom3D.Line3d
            Set oLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(3), dblPnts2(4), dblPnts2(5), dblPnts2(6), dblPnts2(7), dblPnts2(8))
            
            Dim oLine3           As IngrGeom3D.Line3d
            Set oLine3 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(6), dblPnts2(7), dblPnts2(8), dblPnts2(9), dblPnts2(10), dblPnts2(11))
            
            Dim oLine4           As IngrGeom3D.Line3d
            Set oLine4 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(9), dblPnts2(10), dblPnts2(11), dblPnts2(0), dblPnts2(1), dblPnts2(2))
        
            Dim oInnerBoundary           As Collection
            Set oInnerBoundary = New Collection
            
            Dim objHoleF         As IngrGeom3D.ComplexString3d
            
            oInnerBoundary.Add oLine1
            oInnerBoundary.Add oLine2
            oInnerBoundary.Add oLine3
            oInnerBoundary.Add oLine4
            
            Dim Startcurve   As New AutoMath.DPosition
            Startcurve.Set dblPnts2(0), dblPnts2(1), dblPnts2(2)
            
            Set objHoleF = PlaceTrCString(Startcurve, oInnerBoundary)
            ObjOuterPlane.AddBoundary objHoleF
        
                         
            'Set the output
            m_OutputColl.AddOutput "DiffuserStiffener1", ObjOuterPlane
            
            Set ObjOuterPlane = Nothing
            Set objHoleF = Nothing
            Set geomFactory = Nothing
            Set oPos = Nothing
            'Set axis = Nothing
        
            'Insert your code for output 4(Diffuser Baffle -2)
            stPoint.Set DiffuserThickness / 2, 0, 0
            Set ObjDiffuserBaffle2 = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width0, Depth0, Width4, Depth4, _
                            DiffuserThickness / 2, False, 0, PI / 2, 0)
                            
            'Set the output
            m_OutputColl.AddOutput "DiffuserBaffle2", ObjDiffuserBaffle2
            Set ObjDiffuserBaffle2 = Nothing
            
            
            'Insert your code for output 5(Diffuser Baffle -3)
            stPoint.Set DiffuserThickness / 2, 0, 0
        
            Set ObjDiffuserBaffle3 = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width1, Depth1, Width3, Depth3, _
                            DiffuserThickness / 2, False, 0, PI / 2, 0)
                            
            'Set the output
            m_OutputColl.AddOutput "DiffuserBaffle3", ObjDiffuserBaffle3
            Set ObjDiffuserBaffle3 = Nothing
            
            
            'Insert your code for output 5(Diffuser Baffle -4)
        
            stPoint.Set DiffuserThickness / 2, 0, 0
        
            Set ObjDiffuserBaffle4 = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width21, Depth21, Width2, Depth2, _
                            DiffuserThickness / 2, False, 0, PI / 2, 0)
        
            'Set the output
            m_OutputColl.AddOutput "DiffuserBaffle3", ObjDiffuserBaffle4
            Set ObjDiffuserBaffle4 = Nothing
            
        Case 162, 163, 164
            'Place Rectangular Neck
            Dim oRectNeck  As Object
            
            stPoint.Set 0, -Depth / 2, -Width / 2
            If lPartDataBasis = 162 Then
                enPoint.Set 0.5 * DiffuserThickness, Depth / 2, Width / 2
            ElseIf lPartDataBasis = 163 Or lPartDataBasis = 164 Then
                enPoint.Set 0.3 * DiffuserThickness, Depth / 2, Width / 2
            End If
                        
            Set oRectNeck = PlaceBox(m_OutputColl, stPoint, enPoint)
            
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRectNeck
            Set oRectNeck = Nothing
            
            If lPartDataBasis = 163 Or lPartDataBasis = 164 Then
                'Place Transition
                Dim oRectTrans     As Object
                Dim oBasePoint     As AutoMath.DPosition
                Set oBasePoint = New AutoMath.DPosition
                
                oBasePoint.Set 0.7 * DiffuserThickness, 0, 0
                
                Set oRectTrans = PlaceTrapezoid(m_OutputColl, _
                            oBasePoint, 0.8 * BWidth, 0.8 * BDepth, 1.2 * Width, _
                            1.2 * Depth, 0.4 * DiffuserThickness, False, 0, -PI / 2, 0)
                                                             
                m_OutputColl.AddOutput "Transition", oRectTrans
                Set oRectTrans = Nothing
                Set oBasePoint = Nothing
            End If
            
            'Place Rectangular Diffuser
            Dim oRectDiffuser   As Object
            
            If lPartDataBasis = 162 Then
                stPoint.Set 0.5 * DiffuserThickness, -0.4 * BWidth, 0.4 * BDepth
                enPoint.Set 0.9 * DiffuserThickness, 0.4 * BWidth, -0.4 * BDepth
            ElseIf lPartDataBasis = 163 Then
                stPoint.Set 0.7 * DiffuserThickness, -0.4 * BWidth, 0.4 * BDepth
                enPoint.Set 0.9 * DiffuserThickness, 0.4 * BWidth, -0.4 * BDepth
            ElseIf lPartDataBasis = 164 Then
                stPoint.Set 0.7 * DiffuserThickness, -0.4 * BWidth, 0.4 * BDepth
                enPoint.Set DiffuserThickness + DiffuserLength1, 0.4 * BWidth, -0.4 * BDepth
            End If
                
            Set oRectDiffuser = PlaceBox(m_OutputColl, stPoint, enPoint)
                        
            m_OutputColl.AddOutput "Diffuser", oRectDiffuser
            Set oRectDiffuser = Nothing
            
            'Place Ceiling Module
            Dim oCeilingModule  As Object
            If lPartDataBasis = 162 Then
                stPoint.Set 0.9 * DiffuserThickness, -0.5 * BWidth, 0.5 * BDepth
                enPoint.Set DiffuserThickness, 0.5 * BWidth, -0.5 * BDepth
            ElseIf lPartDataBasis = 163 Then
                stPoint.Set 0.9 * DiffuserThickness, -0.5 * BWidth, 0.5 * BDepth
                enPoint.Set DiffuserThickness, 0.5 * BWidth, -0.5 * BDepth
            ElseIf lPartDataBasis = 164 Then
                stPoint.Set 0.9 * DiffuserThickness, -0.5 * BWidth, 0.5 * BDepth
                enPoint.Set DiffuserThickness, 0.5 * BWidth, -0.5 * BDepth
            End If
                            
            Set oCeilingModule = PlaceBox(m_OutputColl, stPoint, enPoint)
            
            m_OutputColl.AddOutput "CeiligModule", oCeilingModule
            Set oCeilingModule = Nothing
    End Select

    '=====================================
    'BUILD HVAC NOZZLE: Diffuser Port
    '=====================================
    Dim NozzleFactory           As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle             As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle                 As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort            As GSCADNozzleEntities.IJDistribPort
    Dim EndPrep                 As Long
    Dim FlowDir                 As DistribFlow
    Dim PortStatus              As DistribPortStatus
    Dim DimBaseOuter            As Boolean
    Dim PortDepth               As Double
    Dim cptOffset               As Double
    Dim pos                     As New AutoMath.DPosition
    Dim dir                     As New AutoMath.DVector
    Dim iPortIndex              As Integer
    Dim CornerRadius            As Double
    Dim FlangeWidth             As Double
    Dim CPos                    As New AutoMath.DPosition
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim NozzleLength As Double, Thickness As Double
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 161 Then
    
        'Set HVAC nozzle parameters
        iPortIndex = 1
        EndPrep = 301
        CornerRadius = 0#
        Thickness = 0#
        NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
        PortDepth = 0#
        cptOffset = 0#
    
        ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
        ' some non-zero value for flange thickness
        ' Flange thickness assumed to be negigible thickness
        FlowDir = DistribFlow_IN
        
    Else
        Dim oHvacPort As IJDHvacPort
        Dim oHvacColl As IJDCollection
        'Set HVAC nozzle parameters
        Set oHvacColl = oPartFclt.GetNozzles()
        For iPortIndex = 1 To oHvacColl.Size
            Set oHvacPort = oHvacColl.Item(iPortIndex)
            EndPrep = oHvacPort.EndPrep
            CornerRadius = 0#
            Thickness = oHvacPort.Thickness
            FlangeWidth = oHvacPort.FlangeWidth
            FlowDir = oHvacPort.FlowDirection
            PortDepth = oHvacPort.PortDepth
            cptOffset = oHvacPort.cptOffset
            
            'NozzleLength Has to be GREATER than NozzleFlangeThickness
            If CmpDblLessThanOrEqualTo(Thickness, LINEAR_TOLERANCE) Then
                Thickness = NEGLIGIBLE_THICKNESS
            End If
            If CmpDblLessThan(NozzleLength, Thickness) Then
                NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
            End If
        Next
        Set oHvacPort = Nothing
        oHvacColl.Clear
        Set oHvacColl = Nothing
    End If

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "DiffuserPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir

    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CPos = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
