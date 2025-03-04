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
    Dim oEllipse As IngrGeom3D.Ellipse3d
    Dim oMajor As AutoMath.DPosition
    Dim dMMRatio As Double
        
    
    Dim dTakeOffLoc As Double
    
    If parHVACShape = FlatOval Then
        Dim oFltOvalCurve As Object
        'Create the non persistent Duct (Header)
        oCenter.Set -(dBWidth + dTapHeight + 2 * INCH), 0, -parDepth / 2
        Set oFltOvalCurve = CreFlatOval(oCenter, parWidth, parDepth, 0)
        Set oHeader = oGeomFactory.Projections3d.CreateByCurve(Nothing, oFltOvalCurve, _
                                                    1, 0, 0, _
                                                    (dBWidth + dTapHeight + 10 * INCH), False)
        
        If CmpDblGreaterthan(dTapTopDepth + 2 * parInsulationThickness, parDepth) Then
            Dim dEllCenY As Double
            Set oMajor = New AutoMath.DPosition
            dEllCenY = (parWidth / 2 - parDepth / 2) + Sqr((parDepth / 2) ^ 2 - (dTapTopDepth / 2) ^ 2)
            oCenter.Set -(dTapHeight / 2), -dEllCenY, -parDepth / 2
            oMajor.Set (dBWidth + dTapHeight + parInsulationThickness) / 2, 0, 0
            dMMRatio = (dTapTopDepth + 2 * parInsulationThickness) / _
                                    (dBWidth + dTapHeight + 2 * parInsulationThickness)
            Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        oCenter.X, oCenter.Y, oCenter.Z, 0, -1, 0, _
                                        oMajor.X, oMajor.Y, oMajor.Z, dMMRatio)
            
            'oCenter.Set 0, -(parWidth / 2 + dTapHeight), -parDepth / 2
            oCenter.Set 0, 0, -parDepth / 2
            oDirVector.Set 0, -1, 0
            Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                                dBDepth + 2 * parInsulationThickness, 0)
            Set oTopCurve = PlaceProjectionTrans(oFltOvalCurve, oDirVector, parWidth / 2 + dTapHeight, True)
            Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                            0, -(parWidth / 2 + dTapHeight), -parDepth / 2, _
                                            0, -1, 0)
            Set oIJSurfaceHeader = oTopCurve
            oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
            Set oIJcurve1 = oIntersectionCurve.Item(1)
            
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                oIJcurve1, oEllipse, True)
        Else
            'Create non persistant object to get the Bottom Curve of the TapBody
            oCenter.Set -(dTapHeight / 2), 0, -parDepth / 2
            oDirVector.Set 0, -1, 0
            Set oFltOvalCurve = CreFlatOvalBranch(oCenter, _
                                    (dBWidth + dTapHeight + 2 * parInsulationThickness), _
                                    (dTapTopDepth + 2 * parInsulationThickness), 0)
            
            Set oTapBody = PlaceProjectionTrans(oFltOvalCurve, oDirVector, parWidth + dTapHeight, True)
            
            'Create non persistant object to get the flat-oval shape at the nozzle
            oCenter.Set 0, 0, -parDepth / 2
            Set oFltOvalCurve = CreFlatOvalBranch(oCenter, dBWidth + 2 * parInsulationThickness, _
                                                dBDepth + 2 * parInsulationThickness, 0)
            Set oTopCurve = PlaceProjectionTrans(oFltOvalCurve, oDirVector, (parWidth / 2 + dTapHeight), True)
            
            Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0, -(parWidth / 2 + dTapHeight), -parDepth / 2, _
                                                    0, -1, 0)
            
            'Curve between TapBody and Header
            Set oIJSurfaceHeader = oTapBody
            oIJSurfaceHeader.Intersect oHeader, oIntersectionCurve, IntersectCode
            Set oIJcurve2 = oIntersectionCurve.Item(1)
    
            'Curve at the of the nozzle
            Set oIJSurfaceHeader = oTopCurve
            oIJSurfaceHeader.Intersect oPlane, oIntersectionCurve, IntersectCode
            Set oIJcurve3 = oIntersectionCurve.Item(1)
    
            'Create Tap Body
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oIJcurve2, oIJcurve3, True)
        End If
        dTakeOffLoc = dTapHeight
        
        'Create the Take Off
        oCenter.Set 0, -(parWidth / 2 + dTakeOffLoc), -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, 1.01 * dBWidth + 2 * parInsulationThickness, _
                                            dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness, 0)
        oDirVector.Set 0, -1, 0
        Set objTakeOff = PlaceProjection(m_OutputColl, oFltOvalCurve, oDirVector, 2 * INCH, True)
        Set oFltOvalCurve = Nothing

        'Create Strip
        oCenter.Set 0, -(parWidth / 2 + dTakeOffLoc - 0.004), -parDepth / 2
        Set oFltOvalCurve = CreFlatOvalBranch(oCenter, 1.06 * dBWidth + 2 * parInsulationThickness, _
                                                dBDepth + 0.06 * dBDepth + 2 * parInsulationThickness, 0)
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
        
        oTopSurPts(0).Set dBWidth / 2 + parInsulationThickness, dBDepth / 2 + parInsulationThickness, dTapHeight
        oTopSurPts(1).Set oTopSurPts(0).X, -oTopSurPts(0).Y, oTopSurPts(0).Z
        oTopSurPts(2).Set -oTopSurPts(0).X, -oTopSurPts(0).Y, oTopSurPts(0).Z
        oTopSurPts(3).Set -oTopSurPts(0).X, oTopSurPts(0).Y, oTopSurPts(0).Z
        
        oBotSurPts(0).Set dBWidth / 2 + parInsulationThickness, dTapTopDepth / 2 + parInsulationThickness, 0
        oBotSurPts(1).Set oBotSurPts(0).X, -oBotSurPts(0).Y, oBotSurPts(0).Z
        oBotSurPts(2).Set -(dBWidth / 2 + dTapHeight + parInsulationThickness), -oBotSurPts(0).Y, oBotSurPts(0).Z
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
        
        dTakeOffLoc = dTapHeight

        'Create Take Off
        oStPoint.Set -(1.01 * dBWidth / 2 + parInsulationThickness), _
                        -(dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness) / 2, _
                        dTakeOffLoc
        oEnPoint.Set (1.01 * dBWidth / 2) + parInsulationThickness, _
                        (dBDepth + 0.01 * dBWidth + 2 * parInsulationThickness) / 2, _
                        dTakeOffLoc + 2 * INCH
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Create Strip
        oStPoint.Set -(1.06 * dBWidth / 2 + parInsulationThickness), _
                        -(dBDepth + 0.06 * dBWidth + 2 * parInsulationThickness) / 2, _
                        dTakeOffLoc - 0.004
        oEnPoint.Set (1.06 * dBWidth / 2 + parInsulationThickness), _
                        (dBDepth + 0.06 * dBWidth + 2 * parInsulationThickness) / 2, _
                            dTakeOffLoc + 0.004
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    ElseIf parHVACShape = 4 Then 'Round
        
        oStPoint.Set -(dBWidth + dTapHeight + 13 * INCH), 0, -parWidth / 2
        oEnPoint.Set (dBWidth + dTapHeight + 13 * INCH), 0, -parWidth / 2
        Set oHeader = PlaceCylinderTrans(oStPoint, oEnPoint, parWidth)
        
        If CmpDblGreaterthan(dTapTopWidth + 2 * parInsulationThickness, parWidth) Then
            'Create the Tap Body
            oCenter.Set -(dTapHeight / 2), 0, -parWidth / 2
            Set oMajor = New AutoMath.DPosition
            oMajor.Set (dBWidth + dTapHeight + 2 * parInsulationThickness) / 2, 0, 0
            dMMRatio = (dTapTopWidth + 2 * parInsulationThickness) / _
                                    (dBWidth + dTapHeight + 2 * parInsulationThickness)
            Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                oCenter.X, oCenter.Y, oCenter.Z, 0, 0, 1, _
                                oMajor.X, oMajor.Y, oMajor.Z, dMMRatio)
            oCenter.Set 0, 0, dTapHeight
            Set oTopCurve = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                            oCenter.X, oCenter.Y, oCenter.Z, 0, 0, 1, (dBWidth + 2 * parInsulationThickness) / 2)
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oEllipse, oTopCurve, True)
        Else
        
            'Create the Tap Body
            Set oMajor = New AutoMath.DPosition
            oCenter.Set -(dTapHeight / 2), 0, -parWidth / 2
            oMajor.Set (dBWidth + dTapHeight) / 2, 0, 0
            dMMRatio = dTapTopWidth / (dBWidth + dTapHeight)
            
            Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                    oCenter.X, oCenter.Y, oCenter.Z, 0, 0, 1, _
                                    oMajor.X, oMajor.Y, oMajor.Z, dMMRatio)
            oDirVector.Set 0, 0, 1
            Set oTapBody = PlaceProjectionTrans(oEllipse, oDirVector, parWidth / 2 + dTapHeight, True)
            
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
                
            Set oIJSurfaceHeader = oHeader
            oIJSurfaceHeader.Intersect oTapBody, oIntersectionCurve, IntersectCode
            Set oIJcurve2 = oIntersectionCurve.Item(1)
            
            Set oIJSurfaceHeader = oPlane
            oIJSurfaceHeader.Intersect oTopCurve, oIntersectionCurve, IntersectCode
            Set oIJcurve3 = oIntersectionCurve.Item(1)
    
            'Create Tap Body
            Set objTapBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                        oIJcurve2, oIJcurve3, True)
        End If
        dTakeOffLoc = dTapHeight
        'Create Take Off
        oStPoint.Set 0, 0, dTakeOffLoc
        oEnPoint.Set 0, 0, dTakeOffLoc + 2 * INCH
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * dBWidth + 2 * parInsulationThickness, True)
        
        'Create Strip
        oStPoint.Set 0, 0, dTakeOffLoc - 0.004
        oEnPoint.Set 0, 0, dTakeOffLoc + 0.004
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * dBWidth + 2 * parInsulationThickness, True)
    
    End If
    
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
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub

End Sub



