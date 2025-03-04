VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   ProgID:         SP3DHRectLoLossTap.HRectLoLossTap
'   Author:         RRK
'   Creation Date:  Tuesday, June 05 2007
'   Description:
'   This symbol is created for Lo-Loss Rectangular Tap of McGill Air flow corporation as per CR-120452.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages

Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private Const E_FAIL = &H80004005
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
      PI = 4 * Atn(1)
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim PortDirection As New AutoMath.DVector
    Dim RadialDirection As New AutoMath.DVector
    Dim iOutput     As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parTapHeight As Double
    Dim parTapExtension As Double
    Dim parAngle As Double
    Dim parHVACShape As Long
    Dim parInsulationThickness As Double
    
' Inputs

    Set oPartFclt = arrayOfInputs(1)
    
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parTapHeight = arrayOfInputs(6)
    parTapExtension = arrayOfInputs(7)
    parAngle = arrayOfInputs(8)
    parHVACShape = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
            
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    'Branch Depth should not be greater than Duct run Depth. If so, equal them
    If Not parHVACShape = 1 Then 'Other than rectangular
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
    End If
    
    iOutput = 0

    Dim dBottomPos As Double
    Dim Inch As Double
    Inch = 0.0254

    If parHVACShape = 4 Then 'Round
        dBottomPos = (parWidth / 2) - Sqr((parWidth / 2) ^ 2 - (parBDepth / 2) ^ 2)
    ElseIf parHVACShape = 1 Then 'Rectangular
        dBottomPos = 0
    ElseIf parHVACShape = 3 Then  'FlatOval
        dBottomPos = (parDepth / 2) - Sqr((parDepth / 2) ^ 2 - (parBDepth / 2) ^ 2)
    End If
    
    Dim dTapLength As Double
    dTapLength = ((parTapHeight - parTapExtension) / Tan(parAngle))

    Dim objCurvesCollection  As Collection
    Set objCurvesCollection = New Collection
    
    Dim CurvePoint1 As New AutoMath.DPosition
    Dim CurvePoint2 As New AutoMath.DPosition
    Dim CurvePoint3 As New AutoMath.DPosition
    Dim CurvePoint4 As New AutoMath.DPosition
    
    CurvePoint1.Set -parBWidth / 2, -parBDepth / 2, -dBottomPos
    CurvePoint2.Set CurvePoint1.x, CurvePoint1.y, 0.001
    CurvePoint3.Set CurvePoint1.x, parBDepth / 2, CurvePoint2.z
    CurvePoint4.Set CurvePoint1.x, CurvePoint3.y, -dBottomPos
        
    'Line 1
    Dim oline1   As IngrGeom3D.Line3d
    Set oline1 = New IngrGeom3D.Line3d
        
    Set oline1 = PlaceTrLine(CurvePoint1, CurvePoint2)
    objCurvesCollection.Add oline1
    Set oline1 = Nothing
    
    'Line 2
    Dim oline2   As IngrGeom3D.Line3d
    Set oline2 = New IngrGeom3D.Line3d

    Set oline2 = PlaceTrLine(CurvePoint2, CurvePoint3)
    
    objCurvesCollection.Add oline2
    Set oline2 = Nothing
    
    'Line 3
    Dim oline3   As IngrGeom3D.Line3d
    Set oline3 = New IngrGeom3D.Line3d

    Set oline3 = PlaceTrLine(CurvePoint3, CurvePoint4)
    
    objCurvesCollection.Add oline3
    Set oline3 = Nothing
    
        
    'Arc
    Dim oArc As Object
    
    Dim oCentPos As AutoMath.DPosition
    Set oCentPos = New AutoMath.DPosition
    
    Dim oNormvec As AutoMath.DVector
    Set oNormvec = New AutoMath.DVector
    
    oNormvec.Set 1, 0, 0
    oCentPos.Set -parBWidth / 2, 0, -parDepth / 2
    
    If parHVACShape = 1 Then
        Set oArc = New IngrGeom3D.Line3d
        Set oArc = PlaceTrLine(CurvePoint4, CurvePoint1)
    Else
        Set oArc = New IngrGeom3D.Arc3d
        Set oArc = PlaceTrArcByCenterNorm(CurvePoint4, CurvePoint1, oCentPos, oNormvec)
    End If
    
    objCurvesCollection.Add oArc
    Set oArc = Nothing

    Dim objTapBodyOutline  As IngrGeom3D.ComplexString3d
    Set objTapBodyOutline = PlaceTrCString(CurvePoint1, objCurvesCollection)

    Dim objTapBody1 As Object
    Dim objTapBody2 As Object
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 1, 0, 0

' Insert your code for output 1(Tap Body 1)
    Set objTapBody1 = PlaceProjection(m_OutputColl, objTapBodyOutline, oAxisVec, parBWidth + dTapLength, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTapBody1
    Set objTapBody1 = Nothing
      
' Insert your code for output 2(Tap Body 2)
    Dim objTapBodyTop As IngrGeom3D.ComplexString3d
    Dim objTapBodyBottom As IngrGeom3D.ComplexString3d
        
    Dim CP As New AutoMath.DPosition
    CP.Set dTapLength / 2, 0, 0
    Set objTapBodyBottom = CreSMRectBranch(CP, parBWidth + dTapLength, parBDepth)
    
    CP.Set 0, 0, parTapHeight - parTapExtension
    Set objTapBodyTop = CreSMRectBranch(CP, parBWidth, parBDepth)
    
    Set objTapBody2 = m_GeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            objTapBodyTop, objTapBodyBottom, True)
    
    Set objTapBodyBottom = Nothing
    Set objTapBodyTop = Nothing
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTapBody2
    Set objTapBody2 = Nothing
  
'Creating Output 3(Flange Body)
    Dim FlangePoints(1) As New AutoMath.DPosition
    Dim objFlangeBody As Object
    Dim objFlangeoutline As Object
    
    If (parHVACShape = 1) Then
        Set objFlangeoutline = New IngrGeom3D.Line3d
        
        FlangePoints(0).Set -parBWidth / 2 - (0.5 * Inch), -parBDepth / 2 - (0.5 * Inch), 0
        FlangePoints(1).Set FlangePoints(0).x, -FlangePoints(0).y, FlangePoints(0).z
        
        Set objFlangeoutline = PlaceTrLine(FlangePoints(0), FlangePoints(1))
        
    Else
        Set objFlangeoutline = New IngrGeom3D.Arc3d
        
        Dim dAngle1 As Double
        Dim dAngle2 As Double
        Dim dTotAngle As Double
        
        If CmpDblEqual(parDepth, parBDepth) Then
            dAngle1 = PI / 2
            dAngle2 = 0
        Else
            dAngle1 = Atn(parBDepth / Sqr(parDepth ^ 2 - parBDepth ^ 2))
            dAngle2 = 0.0254 / parDepth
        End If
        
        dTotAngle = dAngle1 + dAngle2
        
        FlangePoints(0).Set -parBWidth / 2 - (0.5 * Inch), -(1.004 * parDepth / 2) * Sin(dTotAngle), (1.004 * parDepth / 2) * Cos(dTotAngle) - parDepth / 2
        FlangePoints(1).Set FlangePoints(0).x, -FlangePoints(0).y, FlangePoints(0).z
        
        oCentPos.Set -parBWidth / 2 - (0.5 * Inch), 0, -parDepth / 2
        oNormvec.Set -1, 0, 0
        
        Set objFlangeoutline = PlaceTrArcByCenterNorm(FlangePoints(0), FlangePoints(1), oCentPos, oNormvec)
    End If
    

    oAxisVec.Set 1, 0, 0
    
    Set objFlangeBody = PlaceProjection(m_OutputColl, objFlangeoutline, oAxisVec, parBWidth + dTapLength + Inch, True)
    
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeBody
    
    Set m_GeomFactory = Nothing
    Set objFlangeBody = Nothing
    Set objFlangeoutline = Nothing
        
' Place Nozzle 1
'=====================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
'=====================================
    Dim lEndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim bDimBaseOuter As Boolean
    Dim dPortDepth As Double
    Dim dCptOffSet As Double
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    Dim dMinHL As Double
    Dim dMinBL As Double
    Dim dNozzleLength As Double
    Dim dFlangeWidth As Double
    Dim dFlangeThickness As Double
    Dim dTakeOffLength As Double
    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep = oHvacPort.EndPrep
        dFlangeThickness = oHvacPort.Thickness
        dFlangeWidth = oHvacPort.FlangeWidth
        FlowDir = oHvacPort.FlowDirection
        dPortDepth = oHvacPort.PortDepth
        dCptOffSet = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dFlangeThickness, LINEAR_TOLERANCE) Then
            dFlangeThickness = 0.0001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing

'Set HVAC nozzle parameters

    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector
    
    oRadialDir.Set 0, 1, 0

' Insert your code for output 4(HvacPort)
    iPortIndex = 1
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, _
                            lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parBWidth, _
                            parBDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle1", dPortDepth, dCptOffSet, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the conenct point of the nozzle
    dNozzleLength = parTapExtension
    
    oPos.Set 0, 0, parTapHeight
    oDir.Set 0, 0, 1

    Set oDistribPort = oHvacNozzle

    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

    oDistribPort.SetDirectionVector oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oRadialDir = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub

End Sub

Private Sub Class_Terminate()
    Set m_GeomFactory = Nothing
End Sub
Public Function CreSMRectBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreSMRectBranch:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.x - HW, CP.y + HD, CP.z
    Pt(2).Set CP.x + HW, CP.y + HD, CP.z
    Pt(3).Set CP.x + HW, CP.y - HD, CP.z
    Pt(4).Set CP.x - HW, CP.y - HD, CP.z

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    
    Set CreSMRectBranch = objCStr
    Set oLine = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    
    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError METHOD
   
End Function


Private Sub ReportUnanticipatedError(InMethod As String)

Const METHOD = "ReportUnanticipatedError:"


    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub




