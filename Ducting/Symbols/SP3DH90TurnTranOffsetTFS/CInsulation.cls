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
'   Copyright (c) 2005-08, Intergraph. All Rights Reserved.
'
'   CInsulation.cls
'   Author:         NN
'   Creation Date:  Sunday, May 29, 2005
'   Description:    HVAC 90 degree Turn Transition with Offset (eccentric, flush with bottom of duct), and Width transitions to Depth;
'                   Created with Skinning Surface: 2 cross sections and 4 traces passing by all points of cross sections.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   29.Mar.2006     kkk             CR-92030  Eccentric Turn Transition
'   07.Sep.2006     dkl             TR-104418 1)Added functionality for implementing parPlaneOfFlip 1.
'                                             2)parPlaneOfFlip/parFlipComponent/parAlignComponent are assigned permissible values
'                                              if user enters invalid values for these parameters
'  08.SEP.2006      KKC             DI-95670  Replace names with initials in all revision history sheets and symbols
'  28.Feb.2008      VRK             TR-136990 RectReducingElbow component placement results in “Error Updating Symbol” for th
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
   Const METHOD = "run"
    On Error GoTo ErrorLabel
    Const NEGLIGIBLE_THICKNESS = 0.0001
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim Width As Double
    Dim Width2 As Double
    Dim Depth As Double
    Dim Depth2 As Double
    Dim ElbowThroatRadius As Double
    Dim parInsulationThickness As Double
    Dim parFlipComponent As Long
    Dim parAlignComponent As Long
    Dim parPlaneOfFlip As Long
    Dim parWidth As Double
    Dim parWidth2 As Double
    Dim parDepth As Double
    Dim parDepth2 As Double
    Dim Origin As IJDPosition
    Set Origin = New DPosition
    Origin.Set 0, 0, 0
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parWidth2 = arrayOfInputs(4)
    parDepth2 = arrayOfInputs(5)
    ElbowThroatRadius = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    parFlipComponent = arrayOfInputs(8)
    parAlignComponent = arrayOfInputs(9)
    parPlaneOfFlip = arrayOfInputs(10)
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim R As Double
    R = ElbowThroatRadius
        
    'If input values for parPlaneOfFlip(0,1)/parFlipComponent(1,2)/parAlignComponent(1,2) are other than those permitted,
    ' then only permisssible values are assigned to them.
    If parFlipComponent > 2 Then parFlipComponent = 2 Else If parFlipComponent < 1 Then parFlipComponent = 1
    If parAlignComponent > 3 Then parAlignComponent = 2 Else If parAlignComponent < 1 Then parAlignComponent = 1
    If parPlaneOfFlip > 1 Then parPlaneOfFlip = 1 Else If parPlaneOfFlip < 0 Then parPlaneOfFlip = 0
    
    If parPlaneOfFlip = 0 Then
        Depth = parDepth
        Width = parWidth
        Depth2 = parDepth2
        Width2 = parWidth2
    Else
        'To produce the Turn trasition for Plane of Turn 1, the values of depth and width are swapped
        'for the 2 end cross-sections. With these values the trace curves(all 4 curves described below) are computed(using the same procedure for PlaneOfFlip=0)
        'This results in surfaces in Plane of turn 0 which are rotated by PI/2 to transform them to Plane of Flip 1.
        'Note however that orientation of Port 1 remains unchanged (hence parWidth and parDepth are used below for its computation)
        Depth = parWidth
        Width = parDepth
        Depth2 = parWidth2
        Width2 = parDepth2
    End If
    'If parInsulationThickness >= R, the radius of the innner arcs(Trace and Trace 4) is zero.
    If CmpDblGreaterthanOrEqualTo(parInsulationThickness, R) Then
        Const C = 0.01
        parInsulationThickness = R - C
    End If
    
    'InsulationThickness is added to all dimensions
    Depth = Depth + parInsulationThickness * 2
    Width = Width + parInsulationThickness * 2
    Depth2 = Depth2 + parInsulationThickness * 2
    Width2 = Width2 + parInsulationThickness * 2
    
    'The center of curvature of the trace curves moves(by parInsulationThickness) because of adding insulation thickness to width and depth of the cross sections.
    'Therefore the radius(R) is reduced by insulation thickness.
    R = R - parInsulationThickness
    
    Dim Hport2 As IJDPosition
    Set Hport2 = New DPosition
    Hport2.Set Origin.x + Width2 / 2 + R, Origin.y + Depth / 2 + R, Origin.z - Width / 2 + Depth2 / 2
    If parFlipComponent = 1 Then
        Hport2.y = -Hport2.y
    End If
    If parAlignComponent = 1 Then
        Hport2.z = -Hport2.z
    ElseIf parAlignComponent = 3 Then
        Hport2.z = Origin.z 'Modifying Hport2 to suit it to parAlignComponent-Centered
    End If
    
'point 1 is the start of cross section at Port1 (Port 1 is located at Origin 0,0,0)
    Dim point1 As IJDPosition
    Set point1 = New DPosition
    point1.Set Origin.x, Origin.y + Depth / 2, Origin.z - Width / 2
    If parFlipComponent = 1 Then
        point1.y = -point1.y
    End If
    If parAlignComponent = 1 Then
       point1.z = -point1.z
    End If
    
'point 2 is the start of cross section at Port2 (Port 2 is located at Hport2)
    Dim Point5 As IJDPosition
    Set Point5 = New DPosition
    If parAlignComponent = 3 Then
        Point5.Set Hport2.x - Width2 / 2, Hport2.y, Origin.z - Depth2 / 2
    Else
        Point5.Set Hport2.x - Width2 / 2, Hport2.y, point1.z
    End If

    Dim iOutput     As Double
    iOutput = 0
        
' Insert your code for outputs 1(TurnTransition)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oCS1 As IngrGeom3D.ComplexString3d
    Dim oCS2 As IngrGeom3D.ComplexString3d
    Dim oTrace As IngrGeom3D.Arc3d
    Dim oTrace1 As IngrGeom3D.BSplineCurve3d
    Dim oTrace2 As IngrGeom3D.BSplineCurve3d
    Dim oTrace3 As IngrGeom3D.BSplineCurve3d
    Dim oTrace4 As IngrGeom3D.BSplineCurve3d

    Dim csObj As IngrGeom3D.IJElements
    Dim trObj As IngrGeom3D.IJElements
    Dim Surfset   As IngrGeom3D.IJElements
    Dim oTurnTransition As Object
    Dim WCaps As Integer
    WCaps = 0

    'Cross Section 1
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Set lines = New Collection
    
    If parFlipComponent = 2 Then
        If parAlignComponent = 2 Or parAlignComponent = 3 Then
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z, point1.x, point1.y - Depth, point1.z)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y - Depth, point1.z, point1.x, point1.y - Depth, point1.z + Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y - Depth, point1.z + Width, point1.x, point1.y, point1.z + Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z + Width, point1.x, point1.y, point1.z)
            lines.Add oLine
        Else
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z, point1.x, point1.y - Depth, point1.z)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y - Depth, point1.z, point1.x, point1.y - Depth, point1.z - Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y - Depth, point1.z - Width, point1.x, point1.y, point1.z - Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z - Width, point1.x, point1.y, point1.z)
            lines.Add oLine
        End If
    Else
        If parAlignComponent = 2 Or parAlignComponent = 3 Then
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z, point1.x, point1.y + Depth, point1.z)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y + Depth, point1.z, point1.x, point1.y + Depth, point1.z + Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y + Depth, point1.z + Width, point1.x, point1.y, point1.z + Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z + Width, point1.x, point1.y, point1.z)
            lines.Add oLine
        Else
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z, point1.x, point1.y + Depth, point1.z)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y + Depth, point1.z, point1.x, point1.y + Depth, point1.z - Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y + Depth, point1.z - Width, point1.x, point1.y, point1.z - Width)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, point1.x, point1.y, point1.z - Width, point1.x, point1.y, point1.z)
            lines.Add oLine
        End If
    End If

    Set oCS1 = PlaceTrCString(point1, lines)
    Set csObj = New JObjectCollection
    csObj.Add oCS1
    
    'Cross Section 2
    Dim lines2           As Collection
    Set lines2 = New Collection
    
    If parAlignComponent = 2 Or parAlignComponent = 3 Then
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x, Point5.y, Point5.z, Point5.x + Width2, Point5.y, Point5.z)
        lines2.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x + Width2, Point5.y, Point5.z, Point5.x + Width2, Point5.y, Point5.z + Depth2)
        lines2.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x + Width2, Point5.y, Point5.z + Depth2, Point5.x, Point5.y, Point5.z + Depth2)
        lines2.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x, Point5.y, Point5.z + Depth2, Point5.x, Point5.y, Point5.z)
        lines2.Add oLine
    Else
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x, Point5.y, Point5.z, Point5.x + Width2, Point5.y, Point5.z)
        lines2.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x + Width2, Point5.y, Point5.z, Point5.x + Width2, Point5.y, Point5.z - Depth2)
        lines2.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x + Width2, Point5.y, Point5.z - Depth2, Point5.x, Point5.y, Point5.z - Depth2)
        lines2.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point5.x, Point5.y, Point5.z - Depth2, Point5.x, Point5.y, Point5.z)
        lines2.Add oLine
    End If
    
    Set oCS2 = PlaceTrCString(Point5, lines2)
    csObj.Add oCS2
    Set trObj = New JObjectCollection

    If parAlignComponent = 1 Or parAlignComponent = 2 Then
        'Trace
        If parFlipComponent = 2 Then
            Set oTrace = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing _
            , point1.x, point1.y + R, point1.z, point1.x, point1.y, point1.z, Point5.x, Point5.y, Point5.z)
        Else
            Set oTrace = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing _
            , point1.x, point1.y - R, point1.z, point1.x, point1.y, point1.z, Point5.x, Point5.y, Point5.z)
        End If
        
        'Trace 2
        Dim points(1 To 8)   As New AutoMath.DPosition
        If parFlipComponent = 2 Then
            points(1).Set point1.x, point1.y - Depth, point1.z
            points(2).Set point1.x + R / 4, point1.y - Depth, point1.z
            points(3).Set point1.x + R / 2, point1.y - Depth, point1.z
            points(4).Set Hport2.x, point1.y - Depth * 0.9, point1.z
            points(5).Set Hport2.x + Width2 / 4, point1.y - Depth * 0.75, point1.z
            points(6).Set Point5.x + Width2, Point5.y - R / 2, Point5.z
            points(7).Set Point5.x + Width2, Point5.y - R / 4, Point5.z
            points(8).Set Point5.x + Width2, Point5.y, Point5.z
        Else
            points(1).Set point1.x, point1.y + Depth, point1.z
            points(2).Set point1.x + R / 4, point1.y + Depth, point1.z
            points(3).Set point1.x + R / 2, point1.y + Depth, point1.z
            points(4).Set Hport2.x, point1.y + Depth * 0.9, point1.z
            points(5).Set Hport2.x + Width2 / 4, point1.y + Depth * 0.75, point1.z
            points(6).Set Point5.x + Width2, Point5.y + R / 2, Point5.z
            points(7).Set Point5.x + Width2, Point5.y + R / 4, Point5.z
            points(8).Set Point5.x + Width2, Point5.y, Point5.z
        End If
        Set oTrace2 = PlaceTrBspline(8, points)
        trObj.Add oTrace
        trObj.Add oTrace2
        
    ElseIf parAlignComponent = 3 Then
        'Trace 1
        Dim pointsT(1 To 8) As New AutoMath.DPosition
        If parFlipComponent = 2 Then
            pointsT(1).Set point1.x, point1.y, point1.z
            pointsT(2).Set point1.x + R / 4, point1.y, point1.z
            pointsT(3).Set point1.x + R / 2, point1.y, point1.z
            pointsT(4).Set Origin.x + R - R / 2, point1.y + R / 4, Origin.z - Width / 3
            pointsT(5).Set Origin.x + R - R / 4, point1.y + R / 2, Origin.z - Width / 3
            pointsT(6).Set Point5.x, Point5.y - R / 2, Point5.z
            pointsT(7).Set Point5.x, Point5.y - R / 4, Point5.z
            pointsT(8).Set Point5.x, Point5.y, Point5.z
        Else
            pointsT(1).Set point1.x, point1.y, point1.z
            pointsT(2).Set point1.x + R / 4, point1.y, point1.z
            pointsT(3).Set point1.x + R / 2, point1.y, point1.z
            pointsT(4).Set Origin.x + R - R / 2, point1.y - R / 4, Origin.z - Width / 3
            pointsT(5).Set Origin.x + R - R / 4, point1.y - R / 2, Origin.z - Width / 3
            pointsT(6).Set Point5.x, Point5.y + R / 2, Point5.z
            pointsT(7).Set Point5.x, Point5.y + R / 4, Point5.z
            pointsT(8).Set Point5.x, Point5.y, Point5.z
        End If
        Set oTrace1 = PlaceTrBspline(8, pointsT)
        trObj.Add oTrace1
        
         'Trace 2
        If parFlipComponent = 2 Then
            points(1).Set point1.x, point1.y - Depth, point1.z
            points(2).Set point1.x + R / 4, point1.y - Depth, point1.z
            points(3).Set point1.x + R / 2, point1.y - Depth, point1.z
            points(4).Set Hport2.x, point1.y - Depth * 0.9, Origin.z - Width / 3
            points(5).Set Hport2.x + Width2 / 4, point1.y - Depth * 0.75, Origin.z - Width / 3
            points(6).Set Point5.x + Width2, Point5.y - R / 2, Point5.z
            points(7).Set Point5.x + Width2, Point5.y - R / 4, Point5.z
            points(8).Set Point5.x + Width2, Point5.y, Point5.z
        Else
            points(1).Set point1.x, point1.y + Depth, point1.z
            points(2).Set point1.x + R / 4, point1.y + Depth, point1.z
            points(3).Set point1.x + R / 2, point1.y + Depth, point1.z
            points(4).Set Hport2.x, point1.y + Depth * 0.9, Origin.z - Width / 3
            points(5).Set Hport2.x + Width2 / 4, point1.y + Depth * 0.75, Origin.z - Width / 3
            points(6).Set Point5.x + Width2, Point5.y + R / 2, Point5.z
            points(7).Set Point5.x + Width2, Point5.y + R / 4, Point5.z
            points(8).Set Point5.x + Width2, Point5.y, Point5.z
        End If
        Set oTrace2 = PlaceTrBspline(8, points)
        trObj.Add oTrace2
    End If

    'Trace 3
    Dim pointsT3(1 To 8)   As New AutoMath.DPosition
    If parFlipComponent = 2 Then
        If parAlignComponent = 2 Then
            pointsT3(1).Set point1.x, point1.y - Depth, point1.z + Width
            pointsT3(2).Set point1.x + R / 4, point1.y - Depth, point1.z + Width
            pointsT3(3).Set point1.x + R / 2, point1.y - Depth, point1.z + Width
            pointsT3(4).Set Hport2.x, point1.y - Depth * 0.9, point1.z + ((Width + Depth2) / 2)
            pointsT3(5).Set Hport2.x + Width2 / 4, point1.y - Depth * 0.75, point1.z + ((Width + Depth2) / 2)
            pointsT3(6).Set Point5.x + Width2, Point5.y - R / 2, Point5.z + Depth2
            pointsT3(7).Set Point5.x + Width2, Point5.y - R / 4, Point5.z + Depth2
            pointsT3(8).Set Point5.x + Width2, Point5.y, Point5.z + Depth2
        ElseIf parAlignComponent = 3 Then
            pointsT3(1).Set point1.x, point1.y - Depth, point1.z + Width
            pointsT3(2).Set point1.x + R / 4, point1.y - Depth, point1.z + Width
            pointsT3(3).Set point1.x + R / 2, point1.y - Depth, point1.z + Width
            pointsT3(4).Set Hport2.x, point1.y - Depth * 0.9, Origin.z + Width / 3
            pointsT3(5).Set Hport2.x + Width2 / 4, point1.y - Depth * 0.75, Origin.z + Width / 3
            pointsT3(6).Set Point5.x + Width2, Point5.y - R / 2, Point5.z + Depth2
            pointsT3(7).Set Point5.x + Width2, Point5.y - R / 4, Point5.z + Depth2
            pointsT3(8).Set Point5.x + Width2, Point5.y, Point5.z + Depth2
        Else
            pointsT3(1).Set point1.x, point1.y - Depth, point1.z - Width
            pointsT3(2).Set point1.x + R / 4, point1.y - Depth, point1.z - Width
            pointsT3(3).Set point1.x + R / 2, point1.y - Depth, point1.z - Width
            pointsT3(4).Set Hport2.x, point1.y - Depth * 0.9, point1.z - ((Width + Depth2) / 2)
            pointsT3(5).Set Hport2.x + Width2 / 4, point1.y - Depth * 0.75, point1.z - ((Width + Depth2) / 2)
            pointsT3(6).Set Point5.x + Width2, Point5.y - R / 2, Point5.z - Depth2
            pointsT3(7).Set Point5.x + Width2, Point5.y - R / 4, Point5.z - Depth2
            pointsT3(8).Set Point5.x + Width2, Point5.y, Point5.z - Depth2
        End If
    Else
        If parAlignComponent = 2 Then
            pointsT3(1).Set point1.x, point1.y + Depth, point1.z + Width
            pointsT3(2).Set point1.x + R / 4, point1.y + Depth, point1.z + Width
            pointsT3(3).Set point1.x + R / 2, point1.y + Depth, point1.z + Width
            pointsT3(4).Set Hport2.x, point1.y + Depth * 0.9, point1.z + ((Width + Depth2) / 2)
            pointsT3(5).Set Hport2.x + Width2 / 4, point1.y + Depth * 0.75, point1.z + ((Width + Depth2) / 2)
            pointsT3(6).Set Point5.x + Width2, Point5.y + R / 2, Point5.z + Depth2
            pointsT3(7).Set Point5.x + Width2, Point5.y + R / 4, Point5.z + Depth2
            pointsT3(8).Set Point5.x + Width2, Point5.y, Point5.z + Depth2
        ElseIf parAlignComponent = 3 Then
            pointsT3(1).Set point1.x, point1.y + Depth, point1.z + Width
            pointsT3(2).Set point1.x + R / 4, point1.y + Depth, point1.z + Width
            pointsT3(3).Set point1.x + R / 2, point1.y + Depth, point1.z + Width
            pointsT3(4).Set Hport2.x, point1.y + Depth * 0.9, Origin.z + Width / 3
            pointsT3(5).Set Hport2.x + Width2 / 4, point1.y + Depth * 0.75, Origin.z + Width / 3
            pointsT3(6).Set Point5.x + Width2, Point5.y + R / 2, Point5.z + Depth2
            pointsT3(7).Set Point5.x + Width2, Point5.y + R / 4, Point5.z + Depth2
            pointsT3(8).Set Point5.x + Width2, Point5.y, Point5.z + Depth2
        Else
            pointsT3(1).Set point1.x, point1.y + Depth, point1.z - Width
            pointsT3(2).Set point1.x + R / 4, point1.y + Depth, point1.z - Width
            pointsT3(3).Set point1.x + R / 2, point1.y + Depth, point1.z - Width
            pointsT3(4).Set Hport2.x, point1.y + Depth * 0.9, point1.z - ((Width + Depth2) / 2)
            pointsT3(5).Set Hport2.x + Width2 / 4, point1.y + Depth * 0.75, point1.z - ((Width + Depth2) / 2)
            pointsT3(6).Set Point5.x + Width2, Point5.y + R / 2, Point5.z - Depth2
            pointsT3(7).Set Point5.x + Width2, Point5.y + R / 4, Point5.z - Depth2
            pointsT3(8).Set Point5.x + Width2, Point5.y, Point5.z - Depth2
        End If
    End If
    Set oTrace3 = PlaceTrBspline(8, pointsT3)
    trObj.Add oTrace3
  
    'TR-136990
    'Trace 4
    Dim pointsT4(1 To 8)   As New AutoMath.DPosition
    If parFlipComponent = 2 Then
        If parAlignComponent = 2 Then
            pointsT4(1).Set point1.x, point1.y, point1.z + Width
            pointsT4(2).Set point1.x + R / 4, point1.y, point1.z + Width
            pointsT4(3).Set point1.x + R / 2, point1.y, point1.z + Width
            pointsT4(4).Set Origin.x - R / 2 + R, point1.y + R / 10, point1.z + ((Width + Depth2) / 2)
            pointsT4(5).Set Origin.x - R / 4 + R, point1.y + R / 7, point1.z + ((Width + Depth2) / 2)
            pointsT4(6).Set Origin.x + R, Point5.y - R / 2, Point5.z + Depth2
            pointsT4(7).Set Origin.x + R, Point5.y - R / 4, Point5.z + Depth2
            pointsT4(8).Set Origin.x + R, Point5.y, Point5.z + Depth2
        ElseIf parAlignComponent = 3 Then
            pointsT4(1).Set point1.x, point1.y, point1.z + Width
            pointsT4(2).Set point1.x + R / 4, point1.y, point1.z + Width
            pointsT4(3).Set point1.x + R / 2, point1.y, point1.z + Width
            pointsT4(4).Set Origin.x - R / 2 + R, point1.y + R / 4, Origin.z + Width / 3 'point1.z + ((Width + Depth2) / 2)
            pointsT4(5).Set Origin.x - R / 4 + R, point1.y + R / 2, Origin.z + Width / 3   'point1.z + ((Width + Depth2) / 2)
            pointsT4(6).Set Origin.x + R, Point5.y - R / 2, Point5.z + Depth2
            pointsT4(7).Set Origin.x + R, Point5.y - R / 4, Point5.z + Depth2
            pointsT4(8).Set Origin.x + R, Point5.y, Point5.z + Depth2
        Else
            pointsT4(1).Set point1.x, point1.y, point1.z - Width
            pointsT4(2).Set point1.x + R / 4, point1.y, point1.z - Width
            pointsT4(3).Set point1.x + R / 2, point1.y, point1.z - Width
            pointsT4(4).Set Origin.x - R / 2 + R, point1.y + R / 10, point1.z - ((Width + Depth2) / 2)
            pointsT4(5).Set Origin.x - R / 4 + R, point1.y + R / 7, point1.z - ((Width + Depth2) / 2)
            pointsT4(6).Set Origin.x + R, Point5.y - R / 2, Point5.z - Depth2
            pointsT4(7).Set Origin.x + R, Point5.y - R / 4, Point5.z - Depth2
            pointsT4(8).Set Origin.x + R, Point5.y, Point5.z - Depth2
        End If
    Else
        If parAlignComponent = 2 Then
            pointsT4(1).Set point1.x, point1.y, point1.z + Width
            pointsT4(2).Set point1.x + R / 4, point1.y, point1.z + Width
            pointsT4(3).Set point1.x + R / 2, point1.y, point1.z + Width
            pointsT4(4).Set Origin.x - R / 2 + R, point1.y - R / 10, point1.z + ((Width + Depth2) / 2)
            pointsT4(5).Set Origin.x - R / 4 + R, point1.y - R / 7, point1.z + ((Width + Depth2) / 2)
            pointsT4(6).Set Origin.x + R, Point5.y + R / 2, Point5.z + Depth2
            pointsT4(7).Set Origin.x + R, Point5.y + R / 4, Point5.z + Depth2
            pointsT4(8).Set Origin.x + R, Point5.y, Point5.z + Depth2
        ElseIf parAlignComponent = 3 Then
            pointsT4(1).Set point1.x, point1.y, point1.z + Width
            pointsT4(2).Set point1.x + R / 4, point1.y, point1.z + Width
            pointsT4(3).Set point1.x + R / 2, point1.y, point1.z + Width
            pointsT4(4).Set Origin.x - R / 2 + R, point1.y - R / 4, Origin.z + Width / 3 'point1.z + ((Width + Depth2) / 2)
            pointsT4(5).Set Origin.x - R / 4 + R, point1.y - R / 2, Origin.z + Width / 3   'point1.z + ((Width + Depth2) / 2)
            pointsT4(6).Set Origin.x + R, Point5.y + R / 2, Point5.z + Depth2
            pointsT4(7).Set Origin.x + R, Point5.y + R / 4, Point5.z + Depth2
            pointsT4(8).Set Origin.x + R, Point5.y, Point5.z + Depth2
        Else
            pointsT4(1).Set point1.x, point1.y, point1.z - Width
            pointsT4(2).Set point1.x + R / 4, point1.y, point1.z - Width
            pointsT4(3).Set point1.x + R / 2, point1.y, point1.z - Width
            pointsT4(4).Set Origin.x - R / 2 + R, point1.y - R / 10, point1.z - ((Width + Depth2) / 2)
            pointsT4(5).Set Origin.x - R / 4 + R, point1.y - R / 7, point1.z - ((Width + Depth2) / 2)
            pointsT4(6).Set Origin.x + R, Point5.y + R / 2, Point5.z - Depth2
            pointsT4(7).Set Origin.x + R, Point5.y + R / 4, Point5.z - Depth2
            pointsT4(8).Set Origin.x + R, Point5.y, Point5.z - Depth2
        End If
    End If
    Set oTrace4 = PlaceTrBspline(8, pointsT4)
    trObj.Add oTrace4
    
'''
'''''      'Trace 4
'''''    If parFlipComponent = 2 Then
'''''        If parAlignComponent = 2 Then
'''''            Set oTrace4 = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, point1.x - R, point1.y + R * 2, point1.z + ((Width + Depth2) / 2), point1.x, point1.y, point1.z + Width, Point5.x, Point5.y, Point5.z + Depth2)
'''''        Else
'''''            Set oTrace4 = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, point1.x - R, point1.y + R * 2, point1.z - ((Width + Depth2) / 2), point1.x, point1.y, point1.z - Width, Point5.x, Point5.y, Point5.z - Depth2)
'''''        End If
'''''    Else
'''''        If parAlignComponent = 2 Then
'''''            Set oTrace4 = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, point1.x - R, point1.y - R * 2, point1.z + ((Width + Depth2) / 2), point1.x, point1.y, point1.z + Width, Point5.x, Point5.y, Point5.z + Depth2)
'''''        Else
'''''            Set oTrace4 = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, point1.x - R, point1.y - R * 2, point1.z - ((Width + Depth2) / 2), point1.x, point1.y, point1.z - Width, Point5.x, Point5.y, Point5.z - Depth2)
'''''        End If
'''''    End If
'''''
'''''    trObj.Add oTrace4

    'Create Skinning Surface with 4 Traces and 2 Cross sections
    Set Surfset = oGeomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
     trObj, csObj, WCaps)
    ' Set the output
    For Each oTurnTransition In Surfset
        If parPlaneOfFlip = 1 Then
            'The surfaces computed above by skinning correspond to plane when Plane of Flip is 0
            'Therefore they are rotated by PI/2 X axis.
            Dim dAngle As Double
            Dim oTransMat As IJDT4x4
            Set oTransMat = New AutoMath.DT4x4
            Dim oAboutVec As IJDVector
            If oAboutVec Is Nothing Then Set oAboutVec = New DVector
            oAboutVec.Set 1, 0, 0
            oTransMat.LoadIdentity
            dAngle = PI / 2
            oTransMat.Rotate dAngle, oAboutVec
            oTurnTransition.Transform oTransMat
        End If
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTurnTransition
    Next oTurnTransition
    
    Set oTurnTransition = Nothing
    Set Surfset = Nothing
    Set csObj = Nothing
    Set trObj = Nothing
    Set oGeomFactory = Nothing
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    For iCount = 1 To lines2.Count
        lines2.Remove 1
    Next iCount
    Set lines = Nothing
    Set lines2 = Nothing
    Set oCS1 = Nothing
    Set oCS2 = Nothing
    Set oTrace = Nothing
    Set oTrace2 = Nothing
    Set oTrace3 = Nothing
    Set oTrace4 = Nothing
    Set Hport2 = Nothing
    Set point1 = Nothing
    Set Point5 = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

