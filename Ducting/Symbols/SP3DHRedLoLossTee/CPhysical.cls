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
'   Author:          RUK
'   Creation Date:  Thursday, June 7 2007
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const INCH = 0.0254
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parHvacShape As Integer
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parWidth2 As Double
    Dim parDepth2 As Double
    Dim parBDepth As Double
    Dim parBWidth As Double
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHvacShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parWidth2 = arrayOfInputs(5)
    parDepth2 = arrayOfInputs(6)
    parBWidth = arrayOfInputs(7)
    parBDepth = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    iOutput = 0
    
    'Origin is taken at the mid point of the Header
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Set the Default Shape as Round
    If Not parHvacShape = 4 And Not parHvacShape = Rectangular Then
        parHvacShape = 4
    End If
    'Check the Conditions for parBWidth and dBranchLength, and determine the values
'    Condition 1:
'            a)  parBWidth <= parWidth (in case of Round)
'            b)  parBDepth <= parDepth (in case of Rectangular)
'    Condition 2:
'            a)  parBWidth (Round) or parBDepth (Rectangular)            dBranchLength
'                          (Inches)                                        (Inches)
'                            3 - 8                                               4
'                            8.5 - 14                                            7
'                            14.5 - 26                                           10
'                            27 or Larger                                        13
        
    Dim dBranchLength As Double
    Dim dBranchDepth As Double
    Dim dBranchWidth As Double
    
    If parHvacShape = 4 Then 'Round
        'Condition 1
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            dBranchWidth = parWidth
        Else
            dBranchWidth = parBWidth
        End If
        'Condition 2
        If CmpDblGreaterthan(parBWidth, 3 * INCH) And CmpDblLessThan(parBWidth, 8 * INCH) Then
            dBranchLength = 4 * INCH
        ElseIf CmpDblGreaterthan(parBWidth, 8.5 * INCH) And CmpDblLessThan(parBWidth, 14 * INCH) Then
            dBranchLength = 7 * INCH
        ElseIf CmpDblGreaterthan(parBWidth, 14.5 * INCH) And CmpDblLessThan(parBWidth, 26 * INCH) Then
            dBranchLength = 10 * INCH
        ElseIf CmpDblGreaterthan(parBWidth, 27 * INCH) Then
            dBranchLength = 13 * INCH
        End If
    ElseIf parHvacShape = Rectangular Then
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            dBranchDepth = parDepth
        Else
            dBranchDepth = parBDepth
        End If
        'Condition 2
        If CmpDblGreaterthan(parBDepth, 3 * INCH) And CmpDblLessThan(parBDepth, 8 * INCH) Then
            dBranchLength = 4 * INCH
        ElseIf CmpDblGreaterthan(parBDepth, 8.5 * INCH) And CmpDblLessThan(parBDepth, 14 * INCH) Then
            dBranchLength = 7 * INCH
        ElseIf CmpDblGreaterthan(parBDepth, 14.5 * INCH) And CmpDblLessThan(parBDepth, 26 * INCH) Then
            dBranchLength = 10 * INCH
        ElseIf CmpDblGreaterthan(parBDepth, 27 * INCH) Then
            dBranchLength = 13 * INCH
        End If
    End If

'Create the Header Part 1
    Dim objHeader As Object
    Dim dHeaderLength1 As Double
    If parHvacShape = 4 Then 'Round
        'Assumed that Header Part 1 length is BWidth + BLength + 2"
        dHeaderLength1 = dBranchWidth + dBranchLength + 2 * INCH
        oStPoint.Set -dHeaderLength1 / 2, 0, 0
        oEnPoint.Set dHeaderLength1 / 2, 0, 0
        Set objHeader = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
    End If
    If parHvacShape = Rectangular Then
        'Assumed that Header Part 1 length is BDepth + Blength + 4"
        dHeaderLength1 = dBranchDepth + dBranchLength + 2 * INCH
        oStPoint.Set -dHeaderLength1 / 2, -parDepth / 2, -parWidth / 2
        oEnPoint.Set dHeaderLength1 / 2, parDepth / 2, parWidth / 2
        Set objHeader = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeader
    Set objHeader = Nothing

'Create the TakeOff 1 Of Header
    Dim objTakeOff As Object
    Dim dTakeOffLen As Double
    dTakeOffLen = 2 * INCH
    If parHvacShape = 4 Then 'Round
        oStPoint.Set -dHeaderLength1 / 2, 0, 0
        oEnPoint.Set -(dHeaderLength1 / 2 + dTakeOffLen), 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set -dHeaderLength1 / 2, -1.01 * parDepth / 2, -(parWidth / 2 + 0.01 * parDepth)
        oEnPoint.Set -(dHeaderLength1 / 2 + dTakeOffLen), 1.01 * parDepth / 2, (parWidth / 2 + 0.01 * parDepth / 2)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing

'Create the Strip 1 Of Header
    Dim objStrip As Object
    If parHvacShape = 4 Then 'Round
        oStPoint.Set -(dHeaderLength1 / 2 - 0.004), 0, 0
        oEnPoint.Set -(dHeaderLength1 / 2 + 0.004), 0, 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set -(dHeaderLength1 / 2 - 0.004), -1.06 * parDepth / 2, -(parWidth / 2 + 0.06 * parDepth / 2)
        oEnPoint.Set -(dHeaderLength1 / 2 + 0.004), 1.06 * parDepth / 2, (parWidth / 2 + 0.06 * parDepth / 2)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
'Create the Header Part 2
    Dim dHeaderLength2 As Double
    If parHvacShape = 4 Then 'Round
        'Assumed that Header Part 2 length is Width - Width2.
        'Header Part 2 Length must be in between 1" and 12"
        dHeaderLength2 = parWidth - parWidth2
        If CmpDblLessThan(dHeaderLength2, 1 * INCH) Then
            dHeaderLength2 = 1 * INCH
        End If
        If CmpDblGreaterthan(dHeaderLength2, 12 * INCH) Then
            dHeaderLength2 = 12 * INCH
        End If
        oStPoint.Set dHeaderLength1 / 2, 0, 0
        oEnPoint.Set oStPoint.X + dHeaderLength2, oStPoint.Y, oStPoint.Z
        Set objHeader = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parWidth / 2, parWidth2 / 2, True)
    End If
    If parHvacShape = Rectangular Then
        'Assumed that Header Part 2 length is Depth - Depth2.
        'Header Part 2 Length must be in between 1" and 12"
        dHeaderLength2 = parDepth - parDepth2
        If CmpDblLessThan(dHeaderLength2, 1 * INCH) Then
            dHeaderLength2 = 1 * INCH
        End If
        If CmpDblGreaterthan(dHeaderLength2, 12 * INCH) Then
            dHeaderLength2 = 12 * INCH
        End If
        oStPoint.Set dHeaderLength1 / 2, 0, 0
        Set objHeader = PlaceTrapezoid(m_OutputColl, oStPoint, parDepth, parWidth, parDepth2, parWidth2, _
                                    dHeaderLength2, True, PI / 2, PI / 2, 0)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeader
    Set objHeader = Nothing
    
'Create the TakeOff 2 Of Header
    dTakeOffLen = 2 * INCH
    If parHvacShape = 4 Then 'Round
    oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2, 0, 0
    oEnPoint.Set oStPoint.X + dTakeOffLen, 0, 0
    Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth2, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2, -1.01 * parDepth2 / 2, -(parWidth2 / 2 + 0.01 * parDepth2)
        oEnPoint.Set oStPoint.X + dTakeOffLen, 1.01 * parDepth2 / 2, (parWidth2 / 2 + 0.01 * parDepth2 / 2)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing

'Create the Strip 2 Of Header
    If parHvacShape = 4 Then 'Round
    oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2 - 0.004, 0, 0
    oEnPoint.Set oStPoint.X + 0.008, 0, 0
    Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth2, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 + dHeaderLength2 - 0.004, -1.06 * parDepth2 / 2, -(parWidth2 / 2 + 0.06 * parDepth2 / 2)
        oEnPoint.Set oStPoint.X + 0.008, 1.06 * parDepth2 / 2, (parWidth2 / 2 + 0.06 * parDepth2 / 2)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
'Create the Branch
    Dim dBranchTopMinorDia As Double      'Used for Round Shape
    Dim dBranchTopDia As Double         'Used for Round Shape
    Dim dBranchBottomDepth As Double    'Used for Rectangular
    Dim dBranchTopDepth As Double       'Used for Rectangular
    Dim dBranchBottomWidth As Double    'Used for Rectangular
    Dim dBranchStartY As Double
    
    If parHvacShape = 4 Then 'Round
        Dim objBranch As Object
        dBranchTopDia = dBranchWidth
        'Assumed that
        'dBranchTopMinorDia = dBranchTopDia + 2" if dBranchTopDia <= parWidth - 2"
        'dBranchTopMinorDia = dBranchTopDia if dBranchTopDia > parWidth - 2"
        If CmpDblLessThanOrEqualTo(dBranchTopDia, parWidth - 2 * INCH) Then
            dBranchTopMinorDia = dBranchTopDia + 2 * INCH
        Else
            dBranchTopMinorDia = dBranchTopDia
        End If
        Dim oCenter As New AutoMath.DPosition
        Dim oMajor As New AutoMath.DPosition
        Dim oTopCircle As IngrGeom3D.Circle3d
        Dim oBotEllipse As IngrGeom3D.Ellipse3d
        Dim dMMRatio As Double
        dBranchStartY = Sqr((parWidth / 2) ^ 2 - (dBranchTopMinorDia / 2) ^ 2)
        oCenter.Set (dHeaderLength1 / 2 - 1 * INCH - dBranchTopDia / 2), _
                    -(parDepth / 2 + dBranchLength), 0
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius( _
                            Nothing, oCenter.X, oCenter.Y, oCenter.Z, _
                            0, 1, 0, dBranchTopDia / 2)
        'From Caliculations we get the Major diameter of the ellipse as
        '                dBranchTopDia + Total Length Of the Branch
        'Total Length of the Branch will be Depth / 2 + dBranchLength - dBranchStartY
        Dim dEllCenX As Double
        Dim dTotBLen As Double
        dTotBLen = parWidth / 2 + dBranchLength - dBranchStartY
        dEllCenX = (dHeaderLength1 / 2 - 1 * INCH) - (dBranchTopDia + dTotBLen) / 2
        oCenter.Set dEllCenX, -dBranchStartY, 0
        oMajor.Set (dBranchTopDia + dTotBLen) / 2, 0, 0
        dMMRatio = (dBranchTopMinorDia) / (dBranchTopDia + dTotBLen)
        Set oBotEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                        oCenter.X, oCenter.Y, oCenter.Z, 0, 1, 0, oMajor.X, oMajor.Y, oMajor.Z, dMMRatio)
        
        Set objBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBotEllipse, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "Branch_", objBranch
        Set objBranch = Nothing
        Set oTopCircle = Nothing
        Set oCenter = Nothing
    End If
    If parHvacShape = Rectangular Then
        Dim objBranchColl As New Collection
        
        dBranchTopDepth = dBranchDepth
        dBranchBottomDepth = dBranchTopDepth + dBranchLength
        'dBranchBottomWidth = dBranchBottomDepth
        'Assumed that
        'dBranchBottomWidth = dBranchTopDepth + 2" if dBranchTopDepth <= parDepth - 2"
        'dBranchBottomWidth = dBranchTopDepth if dBranchTopDepth > parDepth - 2"
        If CmpDblLessThanOrEqualTo(dBranchTopDepth, parDepth - 2 * INCH) Then
            dBranchBottomWidth = dBranchTopDepth + 2 * INCH
        Else
            dBranchBottomWidth = dBranchTopDepth
        End If
        Dim oTopSurPts(0 To 3) As IJDPosition
        Dim oBotSurPts(0 To 3) As IJDPosition
        Dim iCount As Integer
        
        For iCount = 0 To 3
            Set oTopSurPts(iCount) = New DPosition
            Set oBotSurPts(iCount) = New DPosition
        Next iCount
        
        oTopSurPts(0).Set dHeaderLength1 / 2 - 1 * INCH, parDepth / 2 + dBranchLength, parBWidth / 2
        oTopSurPts(1).Set oTopSurPts(0).X, oTopSurPts(0).Y, -oTopSurPts(0).Z
        oTopSurPts(2).Set oTopSurPts(0).X - dBranchTopDepth, oTopSurPts(0).Y, -oTopSurPts(0).Z
        oTopSurPts(3).Set oTopSurPts(2).X, oTopSurPts(0).Y, oTopSurPts(0).Z
        
        oBotSurPts(0).Set oTopSurPts(0).X, parDepth / 2, dBranchBottomWidth / 2
        oBotSurPts(1).Set oBotSurPts(0).X, oBotSurPts(0).Y, -oBotSurPts(0).Z
        oBotSurPts(2).Set oBotSurPts(0).X - dBranchBottomDepth, oBotSurPts(0).Y, -oBotSurPts(0).Z
        oBotSurPts(3).Set oBotSurPts(2).X, oBotSurPts(0).Y, oBotSurPts(0).Z
        Set objBranchColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To objBranchColl.Count
            m_OutputColl.AddOutput "Branch_", objBranchColl(iCount)
        Next
        For iCount = 1 To objBranchColl.Count
            objBranchColl.Remove 1
        Next iCount
        For iCount = 0 To 3
            Set oTopSurPts(iCount) = Nothing
            Set oBotSurPts(iCount) = Nothing
        Next iCount
    End If
    
'Create the TakeOff Of Branch
    If parHvacShape = 4 Then 'Round
        Dim dBTakeOffDia As Double
        dBTakeOffDia = 1.01 * dBranchTopDia
        oStPoint.Set (dHeaderLength1 / 2 - 1 * INCH - dBranchTopDia / 2), _
                    -(parDepth / 2 + dBranchLength), 0
        oEnPoint.Set (dHeaderLength1 / 2 - 1 * INCH - dBranchTopDia / 2), _
                    -(parDepth / 2 + dBranchLength + 2 * INCH), 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBTakeOffDia, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 - 1 * INCH + 0.01 * dBranchTopDepth, parDepth / 2 + dBranchLength, parBWidth / 2 + 0.01 * dBranchTopDepth
        oEnPoint.Set oStPoint.X - 1.02 * dBranchTopDepth, oStPoint.Y + 2 * INCH, -(parBWidth / 2 + 0.01 * dBranchTopDepth)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTakeOff
    Set objTakeOff = Nothing

'Create the Strip Of Branch
    If parHvacShape = 4 Then 'Round
        Dim dBStripDia As Double
        dBStripDia = 1.06 * dBranchTopDia
        oStPoint.Set (dHeaderLength1 / 2 - 1 * INCH - dBranchTopDia / 2), _
                    -(parDepth / 2 + dBranchLength - 0.004), 0
        oEnPoint.Set (dHeaderLength1 / 2 - 1 * INCH - dBranchTopDia / 2), _
                    -(parDepth / 2 + dBranchLength + 0.004), 0
        Set objStrip = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBStripDia, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength1 / 2 - 1 * INCH + 0.06 * dBranchTopDepth, parDepth / 2 + dBranchLength - 0.004, parBWidth / 2 + 0.06 * dBranchTopDepth
        oEnPoint.Set oStPoint.X - 1.12 * dBranchTopDepth, oStPoint.Y + 0.008, -(parBWidth / 2 + 0.06 * dBranchTopDepth)
        Set objStrip = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrip
    Set objStrip = Nothing
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing
        
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
    
    iPortIndex = 1
    If parHvacShape = Rectangular Then
        dDepth = parDepth
        dWidth = parWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHvacShape = 4 Then 'Round
        dDepth = parWidth
        dWidth = parWidth
        lCSType = GSCADNozzleEntities.Round
    End If
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -dHeaderLength1 / 2
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    oDistribPort.SetDirectionVector dir
    If parHvacShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

'==================================================
'BUILD HVACNOZZLE2 at the end of the Reducing Tee
'==================================================
    iPortIndex = 2
    
    If parHvacShape = Rectangular Then
        dDepth = parDepth2
        dWidth = parWidth2
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width2
    If parHvacShape = 4 Then 'Round
        dDepth = parWidth2
        dWidth = parWidth2
        lCSType = GSCADNozzleEntities.Round
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
    
    'Position of the nozzle should be the conenct point of the nozzle

    X = dHeaderLength1 / 2 + dHeaderLength2
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    oDistribPort.SetDirectionVector dir
    
    If parHvacShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
'============================================
'BUILD HVACNOZZLE3 at the end of the Branch
'============================================
    iPortIndex = 3
    If parHvacShape = 4 Then 'Round
        dDepth = dBranchTopDia
        dWidth = dBranchTopDia
        lCSType = GSCADNozzleEntities.Round
    End If
    If parHvacShape = Rectangular Then
        dDepth = dBranchTopDepth
        dWidth = parBWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    X = (dHeaderLength1 / 2 - 1 * INCH - dBranchTopDia / 2 - dBranchTopDepth / 2)
    If parHvacShape = 4 Then 'Round
        Y = -(parWidth / 2 + dBranchLength)
    End If
    If parHvacShape = Rectangular Then
        Y = parDepth / 2 + dBranchLength
    End If
    Z = 0#
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 0, -1, 0
    oDistribPort.SetDirectionVector dir
    
    If parHvacShape = Rectangular Then
        Dim oRadialDir As AutoMath.DVector
        Set oRadialDir = New AutoMath.DVector
        oDir.Set 0, 1, 0
        oRadialDir.Set 1, 0, 0
        oDistribPort.SetDirectionVector oDir
        oDistribPort.SetRadialOrient oRadialDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oDir = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
