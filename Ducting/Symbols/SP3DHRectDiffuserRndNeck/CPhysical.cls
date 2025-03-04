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
'   ProgID:         SP3DHRecDiffuserRndN.RecDiffuserRndN
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.SEP.2008     PK   CR-149521 Modified the symbol code for the partdatabasis cases 165,166,167,168,169,170
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private RAD          As Double
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
    Dim ObjDiffuserRndNeck      As Object
    Dim ObjDiffuserRectNeck     As Object
    Dim ObjDiffuserProjection   As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle2      As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle3      As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle4      As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle5      As IngrGeom3D.RuledSurface3d
    
    Dim Width                   As Double
        
    Dim BWidth                  As Double
    Dim BDepth                  As Double
    
    Dim NeckWidth               As Double
    Dim NeckDepth               As Double
    
    Dim RoundNeckLength         As Double
    Dim RectNeckLength          As Double
    Dim Conelength              As Double
    Dim DiffuserLength          As Double
    Dim BaffleGap               As Double
    Dim oPos                    As New AutoMath.DPosition
    Dim oDir                    As New AutoMath.DVector
    
    Dim dblPnts1(0 To 11)       As Double
    Dim dblPnts2(0 To 11)       As Double
    
    Dim geomFactory             As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim DiffuserLength1         As Double
    Dim CeilingModuleWidth      As Double
    Dim CeilingModuleDepth      As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'Width = arrayOfInputs(2)
    'NeckWidth = arrayOfInputs(3)
    'NeckDepth = arrayOfInputs(4)
    'BWidth = arrayOfInputs(5)
    'BDepth = arrayOfInputs(6)
    'DiffuserLength = arrayOfInputs(7)
    'RoundNeckLength = arrayOfInputs(8)
    'RectNeckLength = arrayOfInputs(9)
    'DiffuserLength1 = arrayOfInputs(10)
    'CeilingModuleWidth = arrayOfInputs(11)
    'CeilingModuleDepth = arrayOfInputs(12)
    
    'Partdatabasis cases implemented:
    'Rectangular diffuser, with round neck, Type A-165
    'Rectangular diffuser, with round neck, Type B-166
    'Rectangular diffuser, with round neck, Type C-167
    'Rectangular diffuser, with round neck, Type D-168
    'Rectangular diffuser, with round neck, Type E-169
    'Rectangular diffuser, with round neck, Type F-170
        
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 165   'Default   or  Rectangular diffuser, with round neck, Type A
            Width = arrayOfInputs(2)
            NeckWidth = arrayOfInputs(3)
            NeckDepth = arrayOfInputs(4)
            'Here BWidth & BDepth are the Diffuser's Width and Depth
            'at the larger side.No port is being routed from this end.
            BWidth = arrayOfInputs(5)
            BDepth = arrayOfInputs(6)
            DiffuserLength = arrayOfInputs(7)
            RoundNeckLength = arrayOfInputs(8)
            RectNeckLength = arrayOfInputs(9)
        
        Case 166            'Rectangular diffuser, with round neck, Type B
            Width = arrayOfInputs(2)
            'Here BWidth & BDepth are the Diffuser's Width and Depth
            'at the larger side.No port is being routed from this end.
            BWidth = arrayOfInputs(5)
            BDepth = arrayOfInputs(6)
            DiffuserLength = arrayOfInputs(7)
            
        Case 167            'Rectangular diffuser, with round neck, Type C
            Width = arrayOfInputs(2)
            'Here BWidth & BDepth are the Diffuser's Width and Depth
            'at the larger side.No port is being routed from this end.
            BWidth = arrayOfInputs(5)
            BDepth = arrayOfInputs(6)
            DiffuserLength = arrayOfInputs(7)
            
        Case 168            'Rectangular diffuser, with round neck, Type D
            Width = arrayOfInputs(2)
            'Here BWidth & BDepth are the Diffuser's Width and Depth
            'at the larger side.No port is being routed from this end.
            BWidth = arrayOfInputs(5)
            BDepth = arrayOfInputs(6)
            DiffuserLength = arrayOfInputs(7)
            DiffuserLength1 = arrayOfInputs(10)
            
         Case 169           'Rectangular diffuser, with round neck, Type E
            Width = arrayOfInputs(2)
            'Here BWidth & BDepth are the Diffuser's Width and Depth
            'at the larger side.No port is being routed from this end.
            BWidth = arrayOfInputs(5)
            BDepth = arrayOfInputs(6)
            DiffuserLength = arrayOfInputs(7)
            RoundNeckLength = arrayOfInputs(8)
            DiffuserLength1 = arrayOfInputs(10)
                        
         Case 170           'Rectangular diffuser, with round neck, Type F
            Width = arrayOfInputs(2)
            'Here BWidth & BDepth are the Diffuser's Width and Depth
            'at the larger side.No port is being routed from this end.
            BWidth = arrayOfInputs(5)
            BDepth = arrayOfInputs(6)
            DiffuserLength = arrayOfInputs(7)
            RoundNeckLength = arrayOfInputs(8)
            DiffuserLength1 = arrayOfInputs(10)
            CeilingModuleWidth = arrayOfInputs(11)
            CeilingModuleDepth = arrayOfInputs(12)
                        
    End Select
    
    Select Case lPartDataBasis
     
        Case Is <= 1, 165
    
            BaffleGap = (25.4 / 1000)
            
            Dim Width5 As Double
            Dim Depth5 As Double
            Dim StripThickness As Double
            
            Conelength = DiffuserLength - RoundNeckLength - RectNeckLength
            
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
               
            Width0 = NeckWidth - BaffleGap * 2
            Depth0 = NeckDepth - BaffleGap * 2
            
            Width1 = Width0 - BaffleGap * 2
            Depth1 = Depth0 - BaffleGap * 2
            
            Width21 = Width1 - BaffleGap * 2
            Depth21 = Depth1 - BaffleGap * 2
            
            If CmpDblLessThanOrEqualTo(Width21, 0) Or CmpDblLessThanOrEqualTo(Depth21, 0) Then
                Width21 = NeckWidth * 0.01
                Depth21 = NeckDepth * 0.01
            End If
            
            iOutput = 0
            oPos.Set 0, 0, 0
            
            'Insert your code for output 1(Diffuser Neck - Round)
          
            Dim stPoint   As New AutoMath.DPosition
            Dim enPoint   As New AutoMath.DPosition
            
            stPoint.Set 0, 0, 0
            enPoint.Set RoundNeckLength, 0, 0
            
            Set ObjDiffuserRndNeck = PlaceCylinder(m_OutputColl, stPoint, enPoint, Width, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserRndNeck
            Set ObjDiffuserRndNeck = Nothing
            
            
            'Insert your code for output 2(Diffuser Neck - Rectangular)
            stPoint.Set RoundNeckLength, -NeckDepth / 2, -NeckWidth / 2
            enPoint.Set RoundNeckLength + RectNeckLength, NeckDepth / 2, NeckWidth / 2
            
            Set ObjDiffuserRectNeck = PlaceBox(m_OutputColl, stPoint, enPoint)
        
            'Set the output
            m_OutputColl.AddOutput "DiffuserNeck", ObjDiffuserRectNeck
            Set ObjDiffuserRectNeck = Nothing
            
            
            'Insert your code for output 2(Neck Projection)
            
            stPoint.Set RoundNeckLength + RectNeckLength, 0, 0
        
            Set ObjDiffuserProjection = PlaceTrapezoid(m_OutputColl, _
                            stPoint, NeckWidth, NeckDepth, Width5, Depth5, _
                            Conelength, False, 0, PI / 2, 0)
                            
            'Set the output
            m_OutputColl.AddOutput "DiffuserProjection", ObjDiffuserProjection
            Set ObjDiffuserProjection = Nothing
            
            
            'Insert your code for output 3(Diffuser Stiffner -1)
            Dim ObjOuterPlane As IngrGeom3D.Plane3d
             
            oPos.Set DiffuserLength, -BDepth / 2, -BWidth / 2
            
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
            
            oPos.Set DiffuserLength, -Depth5 / 2, -Width5 / 2
            
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
            m_OutputColl.AddOutput "DiffuserStiffener", ObjOuterPlane
            
            Set ObjOuterPlane = Nothing
            Set objHoleF = Nothing
            'Set ObjOuterPlane = Nothing
            Set geomFactory = Nothing
            Set oPos = Nothing
            'Set axis = Nothing
        
        
            'Insert your code for output 4(Diffuser Baffle -2)
            
            stPoint.Set RoundNeckLength + RectNeckLength, 0, 0
            Set ObjDiffuserBaffle2 = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width0, Depth0, Width4, Depth4, _
                            Conelength, False, 0, PI / 2, 0)
                            
            'set the output
            m_OutputColl.AddOutput "DiffuserBaffle2", ObjDiffuserBaffle2
            Set ObjDiffuserBaffle2 = Nothing
            
            
            'Insert your code for output 5(Diffuser Baffle -3)
            
            stPoint.Set RoundNeckLength + RectNeckLength, 0, 0
        
            Set ObjDiffuserBaffle3 = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width1, Depth1, Width3, Depth3, _
                            Conelength, False, 0, PI / 2, 0)
                            
            'Set the output
            m_OutputColl.AddOutput "DiffuserBaffle3", ObjDiffuserBaffle3
            Set ObjDiffuserBaffle3 = Nothing
            
            
            'Insert your code for output 5(Diffuser Baffle -3)
            
            stPoint.Set RoundNeckLength + RectNeckLength, 0, 0
        
            Set ObjDiffuserBaffle4 = PlaceTrapezoid(m_OutputColl, _
                            stPoint, Width21, Depth21, Width2, Depth2, _
                            Conelength, False, 0, PI / 2, 0)
                            
            'Set the output
            m_OutputColl.AddOutput "DiffuserBaffle4", ObjDiffuserBaffle4
            Set ObjDiffuserBaffle4 = Nothing
            
        Case 166, 167, 168, 169, 170
            
            'Place Round Neck
            Dim oRoundNeck  As Object
            
            stPoint.Set 0, 0, 0
            If lPartDataBasis = 166 Then
                'Assumption:Round neck length is assumed to be 0.5 times of Diffuser length
                enPoint.Set 0.5 * DiffuserLength, 0, 0
            ElseIf lPartDataBasis = 167 Or lPartDataBasis = 168 Then
                'Assumption:Round neck length is assumed to be 0.3 times of Diffuser length
                enPoint.Set 0.3 * DiffuserLength, 0, 0
            ElseIf lPartDataBasis = 169 Or lPartDataBasis = 170 Then
                enPoint.Set RoundNeckLength, 0, 0
            End If
                        
            Set oRoundNeck = PlaceCylinder(m_OutputColl, stPoint, enPoint, Width, True)
            
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRoundNeck
            Set oRoundNeck = Nothing
            
            If lPartDataBasis = 167 Or lPartDataBasis = 168 Then
            
                'Place Transition
                Dim oRoundTrans     As Object
                
                'Assumptions:
                'Transition length is assumed to 0.4 times the Diffuser Length.
                'Base Radius of the conical transition is assumed to be 0.4 times the Diffuser Width
                'Top Radius of the conical transition is assumed to be 0.6 times the Width
                stPoint.Set 0.3 * DiffuserLength, 0, 0
                enPoint.Set 0.7 * DiffuserLength, 0, 0
                
                Set oRoundTrans = PlaceCone(m_OutputColl, enPoint, stPoint, _
                        0.4 * BWidth, 0.6 * Width, True)
                        
                m_OutputColl.AddOutput "Transition", oRoundTrans
                Set oRoundTrans = Nothing
            End If
            
            If lPartDataBasis = 166 Or lPartDataBasis = 167 Or lPartDataBasis = 168 Then
            
                'Place Rectangular Diffuser
                Dim oRectDiffuser   As Object
                
                If lPartDataBasis = 166 Then
                    stPoint.Set 0.5 * DiffuserLength, -0.4 * BWidth, 0.4 * BDepth
                    enPoint.Set 0.9 * DiffuserLength, 0.4 * BWidth, -0.4 * BDepth
                ElseIf lPartDataBasis = 167 Then
                    stPoint.Set 0.7 * DiffuserLength, -0.4 * BWidth, 0.4 * BDepth
                    enPoint.Set 0.9 * DiffuserLength, 0.4 * BWidth, -0.4 * BDepth
                ElseIf lPartDataBasis = 168 Then
                    stPoint.Set 0.7 * DiffuserLength, -0.4 * BWidth, 0.4 * BDepth
                    enPoint.Set DiffuserLength + DiffuserLength1, 0.4 * BWidth, -0.4 * BDepth
                End If
                    
                Set oRectDiffuser = PlaceBox(m_OutputColl, stPoint, enPoint)
                            
                m_OutputColl.AddOutput "Diffuser", oRectDiffuser
                Set oRectDiffuser = Nothing
                
                'Place Ceiling Module
                Dim oCeilingModule  As Object
                If lPartDataBasis = 166 Then
                    stPoint.Set 0.9 * DiffuserLength, -0.5 * BWidth, 0.5 * BDepth
                    enPoint.Set DiffuserLength, 0.5 * BWidth, -0.5 * BDepth
                ElseIf lPartDataBasis = 167 Then
                    stPoint.Set 0.9 * DiffuserLength, -0.5 * BWidth, 0.5 * BDepth
                    enPoint.Set DiffuserLength, 0.5 * BWidth, -0.5 * BDepth
                ElseIf lPartDataBasis = 168 Then
                    stPoint.Set 0.9 * DiffuserLength, -0.5 * BWidth, 0.5 * BDepth
                    enPoint.Set DiffuserLength, 0.5 * BWidth, -0.5 * BDepth
                End If
                                
                Set oCeilingModule = PlaceBox(m_OutputColl, stPoint, enPoint)
                
                m_OutputColl.AddOutput "CeiligModule", oCeilingModule
                Set oCeilingModule = Nothing
                
            End If
            
            If lPartDataBasis = 169 Or lPartDataBasis = 170 Then
            
                'Place base plane
                'A square plane is created at the end of the round neck
                'with a width of 1.2 times the duct width
                Dim oPlane As Plane3d
                Dim dPlanePoints(0 To 14) As Double
                dPlanePoints(0) = RoundNeckLength
                dPlanePoints(1) = -0.6 * Width
                dPlanePoints(2) = -0.6 * Width
                
                dPlanePoints(3) = RoundNeckLength
                dPlanePoints(4) = -0.6 * Width
                dPlanePoints(5) = 0.6 * Width
                
                dPlanePoints(6) = RoundNeckLength
                dPlanePoints(7) = 0.6 * Width
                dPlanePoints(8) = 0.6 * Width
                
                dPlanePoints(9) = RoundNeckLength
                dPlanePoints(10) = 0.6 * Width
                dPlanePoints(11) = -0.6 * Width
                
                dPlanePoints(12) = RoundNeckLength
                dPlanePoints(13) = -0.6 * Width
                dPlanePoints(14) = -0.6 * Width
                
                Set oPlane = geomFactory.Planes3d.CreateByPoints( _
                            m_OutputColl.ResourceManager, 5, dPlanePoints)

                m_OutputColl.AddOutput "BasePlane", oPlane
                Set oPlane = Nothing
                            
                'Place diffuser with Trapezoids
                Dim oTrapezoid1             As Object
                Dim oTrapezoid2             As Object
                Dim oTrapezoid3             As Object
                Dim oBasePoint              As AutoMath.DPosition
                Dim dEffectiveDiffLength    As Double
                Set oBasePoint = New AutoMath.DPosition
                
                'Create Trapezoid 1
                dEffectiveDiffLength = DiffuserLength - DiffuserLength1 - RoundNeckLength
                
                If lPartDataBasis = 169 Then
                    oBasePoint.Set 0.75 * dEffectiveDiffLength + RoundNeckLength, 0, 0
                    
                    Set oTrapezoid1 = PlaceTrapezoid(m_OutputColl, oBasePoint, _
                            0.8 * BWidth, 0.8 * BDepth, 1.2 * Width, _
                            1.2 * Width, 0.75 * dEffectiveDiffLength, _
                            False, 0, -PI / 2, 0)
                            
                ElseIf lPartDataBasis = 170 Then
                    oBasePoint.Set 0.5 * dEffectiveDiffLength + RoundNeckLength, 0, 0
                    
                    Set oTrapezoid1 = PlaceTrapezoid(m_OutputColl, oBasePoint, _
                            BWidth, BDepth, 1.2 * Width, 1.2 * Width, _
                            0.5 * dEffectiveDiffLength, False, 0, -PI / 2, 0)
                End If
                                        
                m_OutputColl.AddOutput "Trapezoid1", oTrapezoid1
                Set oTrapezoid1 = Nothing
                
                'Create Trapezoid 2
                
                If lPartDataBasis = 169 Then
                    oBasePoint.Set 0.85 * dEffectiveDiffLength + RoundNeckLength, 0, 0
                
                    Set oTrapezoid2 = PlaceTrapezoid(m_OutputColl, oBasePoint, _
                                    0.6 * BWidth, 0.6 * BDepth, 0.9 * Width, _
                                    0.9 * Width, 0.5 * dEffectiveDiffLength, _
                                    False, 0, -PI / 2, 0)
                                    
                ElseIf lPartDataBasis = 170 Then
                    oBasePoint.Set 0.85 * dEffectiveDiffLength + RoundNeckLength, 0, 0
            
                    Set oTrapezoid2 = PlaceTrapezoid(m_OutputColl, oBasePoint, _
                                    0.7 * BWidth, 0.7 * BDepth, 0.9 * Width, _
                                    0.9 * Width, 0.6 * dEffectiveDiffLength, _
                                    False, 0, -PI / 2, 0)
                End If
                
                m_OutputColl.AddOutput "Trapezoid2", oTrapezoid2
                Set oTrapezoid2 = Nothing
               
                
                'Create Trapezoid 3
                If lPartDataBasis = 169 Then
                
                    oBasePoint.Set DiffuserLength, 0, 0
                    Set oTrapezoid3 = PlaceTrapezoid(m_OutputColl, oBasePoint, _
                                        0.5 * BWidth, 0.5 * BDepth, 0.5 * Width, 0.5 * Width, _
                                        0.25 * dEffectiveDiffLength + DiffuserLength1, False, 0, -PI / 2, 0)
                                        
                ElseIf lPartDataBasis = 170 Then
                
                    oBasePoint.Set DiffuserLength, 0, 0
                    Set oTrapezoid3 = PlaceTrapezoid(m_OutputColl, oBasePoint, _
                                        0.5 * BWidth, 0.5 * BDepth, 0.5 * Width, 0.5 * Width, _
                                        0.25 * dEffectiveDiffLength + DiffuserLength1, False, 0, -PI / 2, 0)
                
                End If
                
                m_OutputColl.AddOutput "Trapezoid3", oTrapezoid3
                Set oTrapezoid3 = Nothing
                
                If lPartDataBasis = 169 Or lPartDataBasis = 170 Then
                    
                    'Create Ruled Surface
                    Dim oSurface    As Object
                    Dim iNum        As Integer
                    For iNum = 0 To 1
                        Dim oSquare1         As IngrGeom3D.ComplexString3d
                        Dim oSquare2         As IngrGeom3D.ComplexString3d
                        Dim lines            As Collection
                        Dim oLine            As Object
                        Dim dSurfaceLength   As Double
                        Dim oStPoint         As AutoMath.DPosition
                        Dim dLenFactor1      As Double
                        Dim dLenFactor2      As Double
                        Dim dDiffWidth       As Double
                        Dim dDiffDepth       As Double
                        Dim dLenFactor3      As Double
                                                                       
                        Set lines = New Collection
                        Set oStPoint = New AutoMath.DPosition
                                           
                        If lPartDataBasis = 169 Then
                            dLenFactor1 = 0.75
                            dSurfaceLength = iNum * 0.2 * dEffectiveDiffLength
                            dLenFactor2 = 0.4
                            dLenFactor3 = 0.05
                            dDiffWidth = BWidth
                            dDiffDepth = BDepth
                            
                        ElseIf lPartDataBasis = 170 Then
                            dLenFactor1 = 0.5
                            dSurfaceLength = iNum * 0.35 * dEffectiveDiffLength
                            dLenFactor2 = 0.5
                            dLenFactor3 = 0.15
                            dDiffWidth = CeilingModuleWidth
                            dDiffDepth = CeilingModuleDepth
                        End If
                        
                        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        -dLenFactor2 * BWidth, _
                                        -dLenFactor2 * BDepth, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        dLenFactor2 * BWidth, _
                                        -dLenFactor2 * BDepth)
                        lines.Add oLine
                        
                        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        dLenFactor2 * BWidth, _
                                        -dLenFactor2 * BDepth, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        dLenFactor2 * BWidth, _
                                        dLenFactor2 * BDepth)
                        lines.Add oLine
                        
                        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        dLenFactor2 * BWidth, _
                                        dLenFactor2 * BDepth, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        -dLenFactor2 * BWidth, _
                                        dLenFactor2 * BDepth)
                        lines.Add oLine
                        
                        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        -dLenFactor2 * BWidth, _
                                        dLenFactor2 * BDepth, _
                                        dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        -dLenFactor2 * BWidth, _
                                        -dLenFactor2 * BDepth)
                        lines.Add oLine
                        oStPoint.Set dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength, _
                                        -dLenFactor2 * BWidth, _
                                        -dLenFactor2 * BDepth
                        
                        If iNum = 0 Then
                            Set oSquare1 = PlaceTrCString(oStPoint, lines)
                        Else
                            Set oSquare2 = PlaceTrCString(oStPoint, lines)
                        End If
                    Next
                    
                    Set oSurface = geomFactory.RuledSurfaces3d.CreateByCurves( _
                                m_OutputColl.ResourceManager, oSquare1, oSquare2, False)
                                                
                    m_OutputColl.AddOutput "Surface", oSurface
                    Set oSurface = Nothing
                    Set oLine = Nothing
                    Set oSquare1 = Nothing
                    
                    'Create Ceiling Module
                    Dim oProjection             As Object
                    Dim oSurfset                As IngrGeom3D.IJElements
                    Dim dModulePoints(0 To 14)  As Double
                    Dim oModulePlate            As IngrGeom3D.Plane3d
                    Dim iOutCount               As Integer

                    dModulePoints(0) = dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength
                    dModulePoints(1) = -0.5 * dDiffWidth
                    dModulePoints(2) = -0.5 * dDiffDepth

                    dModulePoints(3) = dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength
                    dModulePoints(4) = -0.5 * dDiffWidth
                    dModulePoints(5) = 0.5 * dDiffDepth

                    dModulePoints(6) = dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength
                    dModulePoints(7) = 0.5 * dDiffWidth
                    dModulePoints(8) = 0.5 * dDiffDepth

                    dModulePoints(9) = dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength
                    dModulePoints(10) = 0.5 * dDiffWidth
                    dModulePoints(11) = -0.5 * dDiffDepth
                    
                    dModulePoints(12) = dSurfaceLength + dLenFactor1 * dEffectiveDiffLength + RoundNeckLength
                    dModulePoints(13) = -0.5 * dDiffWidth
                    dModulePoints(14) = -0.5 * dDiffDepth
                    
                    Set oModulePlate = geomFactory.Planes3d.CreateByPoints( _
                                        Nothing, 5, dModulePoints)

                    oModulePlate.AddHole oSquare2

                    Set oSurfset = geomFactory.GeometryServices.CreateByProjection( _
                                        m_OutputColl.ResourceManager, oModulePlate, 1, 0, 0, dLenFactor3 * dEffectiveDiffLength, True)
                                        
                    For Each oProjection In oSurfset
                        m_OutputColl.AddOutput "CelingModule" & iOutCount, oProjection
                        iOutCount = iOutCount + 1
                    Next oProjection

                    Set oSurfset = Nothing
                    Set oProjection = Nothing
                    Set oModulePlate = Nothing
                    Set oSquare2 = Nothing
                                
                End If
           End If
    End Select
        
 
    '=====================================
    'BUILD HVAC NOZZLE: Diffuser Port
    '=====================================
    Dim NozzleFactory       As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle         As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim EndPrep             As Long
    Dim FlowDir             As DistribFlow
    Dim PortStatus          As DistribPortStatus
    Dim DimBaseOuter        As Boolean
    Dim PortDepth           As Double
    Dim cptOffset           As Double
    Dim pos                 As New AutoMath.DPosition
    Dim dir                 As New AutoMath.DVector
    Dim iPortIndex          As Integer
    Dim CornerRadius        As Double
    Dim FlangeWidth         As Double
    Dim CPos                As New AutoMath.DPosition
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim NozzleLength As Double, Thickness As Double
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 165 Then
        
        'Set HVAC nozzle parameters
        iPortIndex = 1
        EndPrep = 301
        CornerRadius = 0#
        Thickness = 0#
        NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
        PortDepth = 0#
        cptOffset = 0#
        
        'To construct nozzle as crosssection only, use FlangeWidth of 0, and
        'some non-zero value for flange thickness
        'Flange thickness assumed to be negigible thickness
        FlowDir = DistribFlow_IN

    Else
        Dim oHvacPort As IJDHvacPort
        Dim oHvacColl As IJDCollection
        'Set HVAC nozzle parameters
        Set oHvacColl = oPartFclt.GetNozzles()
        For iPortIndex = 1 To oHvacColl.Size
            Set oHvacPort = oHvacColl.Item(iPortIndex)
            iPortIndex = 1
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
        Next iPortIndex
        oHvacColl.Clear
        Set oHvacColl = Nothing
        Set oHvacPort = Nothing
    End If
                    
    'Depth and Width of crosssection will be the same as diameter of the nozzle which
    'is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Width, CornerRadius, DimBaseOuter, PortStatus, _
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

