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
'   Copyright (c) 2007 Intergraph
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Friday, June 8 2007
'   Description:

'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const INCH = 0.0254
Private PI       As Double
Private Const MODULE = "Insulation:" 'Used for error messages
Private Sub Class_Initialize()
 PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
 Dim oPartFclt       As PartFacelets.IJDPart
    Dim CP As New AutoMath.DPosition
    Dim CV As New AutoMath.DPosition
    Dim NozzleLength As Double
    
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    Dim parPlaneOfBranch As Double
    
    Dim PortDirection As AutoMath.DVector
    Set PortDirection = New AutoMath.DVector
    Dim RadialDirection As AutoMath.DVector
    Set RadialDirection = New AutoMath.DVector

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parBWidth = arrayOfInputs(5)
    parBDepth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    parPlaneOfBranch = arrayOfInputs(8)

    iOutput = 0
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If

    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If
'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    Dim dHlength As Double
    Dim dBLength As Double
        
' Insert your code for output (Body)
    
    If (lPartdatabasis <= 1) Or (lPartdatabasis = 5) Then
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            parBWidth = parWidth
        End If
        
        If CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
        
        dHlength = parBWidth + (2 * INCH)
        
        If parPlaneOfBranch = PI / 2 Then
            dBLength = parDepth / 2 + INCH
        Else
            dBLength = parWidth / 2 + INCH
        End If
        
    ElseIf (lPartdatabasis = 10) Then
        If CmpDblGreaterthan(parBWidth, parWidth - (2 * INCH)) Then
            parBWidth = parWidth - (2 * INCH)
        End If
        
        If CmpDblGreaterthan(parBDepth, parDepth - (2 * INCH)) Then
            parBDepth = parDepth - (2 * INCH)
        End If
        
        dHlength = parBWidth + (4 * INCH)
        
        If parPlaneOfBranch = PI / 2 Then
            dBLength = parDepth / 2 + (4 * INCH)
        Else
            dBLength = parWidth / 2 + (4 * INCH)
        End If
    End If
    
    Dim dBWidth As Double
    Dim dBDepth As Double
    
    Dim dBranchStartY As Double
    Dim dBranchDia As Double
    
    If (lPartdatabasis <= 1) Or (lPartdatabasis = 5) Then
        dBranchDia = parBWidth
        dBWidth = parBWidth
        dBDepth = parBDepth
    ElseIf (lPartdatabasis = 10) Then
        dBranchDia = parBDepth + (1 * INCH)
        dBWidth = parBWidth + (1 * INCH)
        dBDepth = parBDepth + (1 * INCH)
    End If
    
    If CmpDblGreaterthan(dBranchDia, parWidth) Then
        dBranchDia = parWidth
    End If
    
    dBranchStartY = Sqr((parWidth / 2) ^ 2 - (dBranchDia / 2) ^ 2)
    
    
'''
'Place Output 4

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oCurve1 As Object
    Dim oCurve2 As Object
    
    Dim oHeaderCurve As Object
    Dim oBranchCurve As Object
    

    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector
    'create the profile for the sweep
    If parHVACShape = FlatOval Then
        If CmpDblGreaterthan(dBDepth, parDepth) Then
            dBDepth = parDepth
        End If
        
        dBranchStartY = ((parWidth - parDepth) / 2) + ((parDepth / 2) ^ 2 - (dBDepth / 2) ^ 2)
        CP.Set 0, dBranchStartY, 0
        Set oCurve1 = CreFlatOvalBranch(CP, dBWidth + (2 * parInsulationThickness), dBDepth + (2 * parInsulationThickness), parPlaneOfBranch)
        
        CP.Set 0, dBLength + (2 * INCH), 0
        Set oBranchCurve = CreFlatOvalBranch(CP, (1.01 * parBWidth) + (2 * parInsulationThickness), (1.01 * parBDepth) + (2 * parInsulationThickness), parPlaneOfBranch)
        
        CP.Set (-dHlength / 2) - (2 * INCH), 0, 0
        Set oHeaderCurve = CreFlatOval(CP, (1.01 * parWidth) + (2 * parInsulationThickness), (1.01 * parDepth) + (2 * parInsulationThickness), parPlaneOfBranch)

        
    ElseIf parHVACShape = 4 Then 'Round=4
        CP.Set 0, dBranchStartY, 0
        Dir.Set 0, 1, 0
        Set oCurve1 = PlaceTrCircleByCenter(CP, Dir, (dBranchDia / 2) + parInsulationThickness)
        
            
        CP.Set 0, dBLength + (2 * INCH), 0
        Set oBranchCurve = PlaceTrCircleByCenter(CP, Dir, (1.01 * parBWidth / 2) + parInsulationThickness)
        
        CP.Set (-dHlength / 2) - (2 * INCH), 0, 0
        Dir.Set 1, 0, 0
        Set oHeaderCurve = PlaceTrCircleByCenter(CP, Dir, (1.01 * parWidth / 2) + parInsulationThickness)
        
    ElseIf parHVACShape = 1 Then
        
        If (parPlaneOfBranch = PI / 2) Then
            CP.Set 0, (parDepth / 2), 0
        Else
            CP.Set 0, (parWidth / 2), 0
        End If
        Set oCurve1 = CreRectBranch(CP, dBWidth + (2 * parInsulationThickness), dBDepth + (2 * parInsulationThickness), parPlaneOfBranch)
        
        
        CP.Set 0, dBLength + (2 * INCH), 0
        
        Set oBranchCurve = CreRectBranch(CP, (1.01 * parBWidth) + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), parPlaneOfBranch)
         
        CP.Set (-dHlength / 2) - (2 * INCH), 0, 0
        Set oHeaderCurve = CreRectangle(CP, (1.01 * parWidth) + (2 * parInsulationThickness), (1.01 * parDepth) + (2 * parInsulationThickness), parPlaneOfBranch)
    End If
    
        
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    Dim objInsHeader As Object
    
    'Place Header Insulation
    oAxisVec.Set 1, 0, 0
    Set objInsHeader = PlaceProjection(m_OutputColl, oHeaderCurve, oAxisVec, dHlength + (4 * INCH), True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHeader
    Set objInsHeader = Nothing
    

    Dim objInsBranchBody As Object
    Set objInsBranchBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oCurve1, oBranchCurve, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBranchBody
    
    Set objInsBranchBody = Nothing
    
    Set oCurve1 = Nothing
    Set oHeaderCurve = Nothing
    Set oBranchCurve = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub
Public Function CreFlatOval(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x, CP.y - (Width - Depth) / 2, CP.z + Depth / 2
    Pt(2).Set CP.x, CP.y + (Width - Depth) / 2, Pt(1).z
    Pt(3).Set CP.x, CP.y + Width / 2, CP.z
    Pt(4).Set CP.x, Pt(2).y, CP.z - Depth / 2
    Pt(5).Set CP.x, Pt(1).y, Pt(4).z
    Pt(6).Set CP.x, CP.y - Width / 2, CP.z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 1, 0, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreFlatOval = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError METHOD
    
End Function
Public Function CreFlatOvalBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(6)           As New AutoMath.DPosition
    
    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x - (Width - Depth) / 2, CP.y, CP.z + Depth / 2
    Pt(2).Set CP.x + (Width - Depth) / 2, CP.y, CP.z + Depth / 2
    Pt(3).Set CP.x + Width / 2, CP.y, CP.z
    Pt(4).Set CP.x + (Width - Depth) / 2, CP.y, CP.z - Depth / 2
    Pt(5).Set CP.x - (Width - Depth) / 2, CP.y, CP.z - Depth / 2
    Pt(6).Set CP.x - Width / 2, CP.y, CP.z
        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 0, 1, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreFlatOvalBranch = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError METHOD
    
End Function
Public Function CreRectangle(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreRectangle:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.x, CP.y - HW, CP.z + HD
    Pt(2).Set CP.x, CP.y + HW, CP.z + HD
    Pt(3).Set CP.x, CP.y + HW, CP.z - HD
    Pt(4).Set CP.x, CP.y - HW, CP.z - HD

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)
    Dim oDirVector As AutoMath.DVector
    
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 1, 0, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreRectangle = objCStr
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
Public Function CreRectBranch(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal PlaneofBranch As Double) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(4)          As New AutoMath.DPosition
    
    Const METHOD = "CreRectangle:"
    On Error GoTo ErrorHandler
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD              As Double
    Dim HW              As Double
    Dim CR              As Double
    HD = Depth / 2
    HW = Width / 2
    


    Pt(1).Set CP.x - HW, CP.y, CP.z + HD
    Pt(2).Set CP.x + HW, CP.y, CP.z + HD
    Pt(3).Set CP.x + HW, CP.y, CP.z - HD
    Pt(4).Set CP.x - HW, CP.y, CP.z - HD

        
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine


    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
    If (PlaneofBranch = 0) Then
        dRotation = 0
        oDirVector.Set 1, 0, 0
    ElseIf (PlaneofBranch = PI / 2) Then
        dRotation = PI / 2
        oDirVector.Set 0, 1, 0
    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreRectBranch = objCStr
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
'''<{(Circle begin)}>
Public Function PlaceTrCircleByCenter(ByRef centerPoint As AutoMath.DPosition, _
                            ByRef normalVector As AutoMath.DVector, _
                            ByRef Radius As Double) _
                            As IngrGeom3D.Circle3d

''' This function creates transient (non-persistent) circle
''' Example of call:
''' Dim point   As new AutoMath.DPosition
''' Dim normal  As new AutoMath.DVector
''' Dim objCircle  As IngrGeom3D.circle3d
''' point.set 0, 0, 0
''' normal.set 0, 0, 1
''' set objCircle = PlaceTrCircleByCenter(point, normal, 2 )
''' ......... use this object (e.g. to create projection)
''' set objCircle = Nothing


    Const METHOD = "PlaceTrCircleByCenter:"
    On Error GoTo ErrorHandler
        
    Dim oCircle As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the Circle"
    ' Create Circle object
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            centerPoint.x, centerPoint.y, centerPoint.z, _
                            normalVector.x, normalVector.y, normalVector.z, _
                            Radius)
    Set PlaceTrCircleByCenter = oCircle
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError METHOD

End Function

Private Sub ReportUnanticipatedError(InMethod As String)

Const METHOD = "ReportUnanticipatedError:"
'    Dim ern As IMSErrorService.JWellKnownErrorNumbers
'    Dim errorService As IMSErrorLog.IJErrorService
'    Dim oTrader As New Trader
'
'    Set errorService = oTrader.Service("ErrorHandler", "")
'
'    ern = errorService.ReportError(Err.Number, MODULE & ":" & InMethod, "UNANTICIPATED", App)
'
'    Select Case ern
'      Case imsAbortApplication:
'            errorService.TerminateApp
'    End Select
'
'    Set errorService = Nothing
'    Set oTrader = Nothing

    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub



