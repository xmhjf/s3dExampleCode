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
'   Author:         RRK
'   Creation Date:  Tuesday, May 29 2007
'   Description:
'   This symbol is prepared for Rectangular tap of McGill Air flow corporation as per CR-120452
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation: 'Used for error messages
Private m_GeomFactory As IngrGeom3D.GeometryFactory
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
    Dim parBDepth As Double
    Dim parBWidth As Double
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
    
    Dim dInsulationWidth As Double
    Dim dInsulationDepth As Double

    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
        
    'Branch Depth should not be greater than Duct run Depth. If so, equal them
    If Not parHVACShape = 1 Then 'Other than rectangular
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
    End If
    
    dInsulationWidth = parWidth + (2 * parInsulationThickness)
    dInsulationDepth = parDepth + (2 * parInsulationThickness)
    
    iOutput = 0
    
    Dim Inch As Double
    Inch = 0.0254

    Dim dBottomPos As Double
    Dim dInsBottomPos As Double

    If parHVACShape = 4 Then 'Round
        dInsBottomPos = (dInsulationWidth / 2) - Sqr((dInsulationWidth / 2) ^ 2 - ((parBDepth / 2) + parInsulationThickness) ^ 2)
        dBottomPos = (parWidth / 2) - Sqr((parWidth / 2) ^ 2 - (parBDepth / 2) ^ 2)
    ElseIf parHVACShape = 1 Then 'Rectangular
        dBottomPos = 0
        dInsBottomPos = 0
    ElseIf parHVACShape = 3 Then  'FlatOval
        dInsBottomPos = (dInsulationDepth / 2) - Sqr((dInsulationDepth / 2) ^ 2 - ((parBDepth / 2) + parInsulationThickness) ^ 2)
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
    
    CurvePoint1.Set -parBWidth / 2 - parInsulationThickness, -parBDepth / 2 - parInsulationThickness, parInsulationThickness - dInsBottomPos
    CurvePoint2.Set CurvePoint1.x, CurvePoint1.y, parInsulationThickness + 0.001
    CurvePoint3.Set CurvePoint1.x, parBDepth / 2 + parInsulationThickness, CurvePoint2.z
    CurvePoint4.Set CurvePoint1.x, CurvePoint3.y, CurvePoint1.z
        
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
    oCentPos.Set -parBWidth / 2 - parInsulationThickness, 0, -parDepth / 2
    
    If parHVACShape = 1 Then
        Set oArc = New IngrGeom3D.Line3d
        Set oArc = PlaceTrLine(CurvePoint4, CurvePoint1)
    Else
        Set oArc = New IngrGeom3D.Arc3d
        Set oArc = PlaceTrArcByCenterNorm(CurvePoint4, CurvePoint1, oCentPos, oNormvec)
    End If
    
    objCurvesCollection.Add oArc
    Set oArc = Nothing

    Dim objInsTapBodyOutline  As IngrGeom3D.ComplexString3d
    Set objInsTapBodyOutline = PlaceTrCString(CurvePoint1, objCurvesCollection)

    Dim objInsTapBody1 As Object
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 1, 0, 0

'Creating Output 5(Insulation Body 1)
    Set objInsTapBody1 = PlaceProjection(m_OutputColl, objInsTapBodyOutline, oAxisVec, parInsulationThickness + 0.001, True)
    
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTapBody1
    
    Set objInsTapBody1 = Nothing
      
'Creating Output 6(Insulation Body 2)
    Dim objTapBodyTop As IngrGeom3D.ComplexString3d
    Dim objTapBodyBottom As IngrGeom3D.ComplexString3d
    Dim objInsTapBody2 As Object
        
    Dim CP As New AutoMath.DPosition
    CP.Set dTapLength / 2, 0, 0
    Set objTapBodyBottom = CreSMRectBranch(CP, parBWidth + dTapLength + (2 * parInsulationThickness), _
                                                            parBDepth + (2 * parInsulationThickness))
    
    CP.Set 0, 0, parTapHeight - parTapExtension
    Set objTapBodyTop = CreSMRectBranch(CP, parBWidth + (2 * parInsulationThickness), _
                                            parBDepth + (2 * parInsulationThickness))
    
    Set objInsTapBody2 = m_GeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            objTapBodyTop, objTapBodyBottom, True)
    
    Set objTapBodyBottom = Nothing
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTapBody2
    
'Create Output 7(Insulation for Tap extension)
    Dim objInsBox As Object
    
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
'
    Set oAxisVec = New AutoMath.DVector

    oAxisVec.Set 0, 0, 1
    
    Set objInsBox = PlaceProjection(m_OutputColl, objTapBodyTop, oAxisVec, parTapExtension, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBox
    
    Set objInsTapBody2 = Nothing
    Set objTapBodyTop = Nothing
    Set m_GeomFactory = Nothing
    Set objInsBox = Nothing

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








