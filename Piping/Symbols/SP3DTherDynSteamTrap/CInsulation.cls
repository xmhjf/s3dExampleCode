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
'   Author:         KKC
'   Creation Date:  Tuesday, Sep 18 2007
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
   
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim parFaceToFace As Double
    Dim parDiameter1 As Double
    Dim parDiameter2 As Double
    Dim parTrapHeight As Double
    Dim parOffset As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parInsulationThickness As Double
    Dim lPartDataBasis As Long
    Dim parInsulationDiameter As Double
    Dim iCount As Long
    
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Set oPartFclt = arrayOfInputs(1)
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

' Inputs
    iOutput = 0
    iCount = 0
    If lPartDataBasis <= 1 Or lPartDataBasis = 425 Then
        parFaceToFace = arrayOfInputs(2)
        parOffset1 = arrayOfInputs(3)
        parOffset2 = arrayOfInputs(4)
        parDiameter1 = arrayOfInputs(5)
        parDiameter2 = arrayOfInputs(6)
        parInsulationThickness = arrayOfInputs(10)
        
    ElseIf lPartDataBasis = 426 Then
        parFaceToFace = arrayOfInputs(2)
        parOffset1 = arrayOfInputs(3)
        parOffset2 = arrayOfInputs(4)
        parDiameter1 = arrayOfInputs(5)
        parDiameter2 = arrayOfInputs(6)
        parInsulationThickness = arrayOfInputs(10)
        
    ElseIf lPartDataBasis = 427 Then
        parFaceToFace = arrayOfInputs(2)
        parOffset = arrayOfInputs(8)
        parTrapHeight = arrayOfInputs(7)
        parInsulationThickness = arrayOfInputs(10)
        
    ElseIf lPartDataBasis = 428 Then
        parFaceToFace = arrayOfInputs(2)
        parOffset1 = arrayOfInputs(3)
        parOffset2 = arrayOfInputs(4)
        parPipingSpecialtyWidth = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(10)
        
    End If
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 425 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        Dim dProjLength As Double
        Dim dAngle As Double
        dProjLength = 0.7 * (parFaceToFace - 2 * flangeThick)
        dAngle = Atn(parOffset / dProjLength)
        
        If CmpDblEqual(parDiameter1, 0) Then parDiameter1 = 0.25 * parFaceToFace
        If CmpDblEqual(parDiameter2, 0) Then parDiameter2 = parDiameter1
        
    ' Insert your code for Output 1 (Cylinder)
        oStPoint.Set -parFaceToFace / 2, 0, 0
        oEnPoint.Set parFaceToFace / 2, 0, 0
        Dim objCylinder As Object
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, 0, 0
        
        Dim objCylinder1 As Object
        Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * pipeDiam + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder1
        Set objCylinder1 = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
       
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, -(parOffset2 - parDiameter2 * Cos(dAngle)), 0
        
        Dim objInclCylinder As Object
        Set objInclCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter2 + 2 * parInsulationThickness, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInclCylinder
        Set objInclCylinder = Nothing
    ' Insert your code for the Output 4 (Hexagonal Head)
         Dim dLineStrPoints(0 To 20)  As Double
         Dim oLineString As IngrGeom3D.LineString3d
         Set oLineString = New IngrGeom3D.LineString3d
         Dim dAngle1 As Double
         dAngle1 = PI / 3
         oStPoint.Set 0, 0.7 * pipeDiam, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parDiameter1 / 2 + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parDiameter1 / 2 + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        Dim oAxisVect As AutoMath.DVector
        Dim ObjHexagonalHead As Object
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset1 - 0.75 * pipeDiam + parInsulationThickness, True)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oAxisVect = Nothing
        
    ElseIf lPartDataBasis = 426 Then
    
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        dProjLength = 0.5 * (parFaceToFace - 2 * flangeThick)
        dAngle = Atn(parOffset / dProjLength)
        If CmpDblEqual(parDiameter1, 0) Then parDiameter1 = 0.35 * parFaceToFace
        If CmpDblEqual(parDiameter2, 0) Then parDiameter2 = 0.6 * parDiameter1
    ' Insert your code for Output 1 (Cylinder)
        oStPoint.Set -parFaceToFace / 2, 0, 0
        oEnPoint.Set parFaceToFace / 2, 0, 0
      
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, 0, 0
        
        Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder1
        Set objCylinder1 = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
       
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, -(parOffset2 - (parDiameter2 / 2 * Cos(dAngle))), 0
        
        Set objInclCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter2 + 2 * parInsulationThickness, True)
        
    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", objInclCylinder
        Set objInclCylinder = Nothing
    ' Insert your code for the Output 4 ( Hexagonal Head)
        
         dAngle1 = PI / 3
         oStPoint.Set 0, 0.7 * pipeDiam, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parDiameter1 / 2 + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parDiameter1 / 2 + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
       
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset1 - 0.6 * pipeDiam + parInsulationThickness, True)

    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
        Set oLineString = Nothing
        Dim ObjHexagonalNut As Object
        Dim ObjHexagonalNut1 As Object
        
        oStPoint.Set -dProjLength - 0.1 * parFaceToFace - parInsulationThickness, 0, 0
    '   First point is considered on positive Z-axis at an angle (PI / 6 + Angle) from positive Y-axis.
        iCount = 0
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x
            dLineStrPoints(3 * iCount - 2) = oStPoint.y + (0.8 * pipeDiam + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (0.8 * pipeDiam + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        
        oAxisVect.Set 1, 0, 0
        
        Set ObjHexagonalNut = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.1 * parFaceToFace + 2 * parInsulationThickness, True)
        Set ObjHexagonalNut1 = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.1 * parFaceToFace + 2 * parInsulationThickness, True)
    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalNut
        Set ObjHexagonalNut = Nothing
        
        Dim oTransMatrix As AutoMath.DT4x4
        Set oTransMatrix = New AutoMath.DT4x4
        
        oTransMatrix.LoadIdentity
        Dim oDir As AutoMath.DVector
        Set oDir = New AutoMath.DVector
        
        oDir.Set dProjLength + 0.1 * parFaceToFace + 2 * parInsulationThickness, 0, 0
        oTransMatrix.Translate oDir
        ObjHexagonalNut1.Transform oTransMatrix
        
     '   Set the output
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalNut1
        Set ObjHexagonalNut1 = Nothing
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oAxisVect = Nothing
    
    ElseIf lPartDataBasis = 427 Then
       RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    '  Insert your code for the bottom Hexagonal head
         dAngle1 = PI / 3
         oStPoint.Set 0, -parOffset - parInsulationThickness, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + oStPoint.x + (parFaceToFace * Tan(PI / 6) + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parFaceToFace * Tan(PI / 6) + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New IngrGeom3D.LineString3d
        Set oAxisVect = New AutoMath.DVector
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset + flangeDiam / 2, True)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
        
    '  Inset your code for the Output 2 (Cylinder)
        oStPoint.Set 0, flangeDiam / 2, 0
        oEnPoint.Set 0, flangeDiam / 2 + 0.1 * parTrapHeight, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.7 * parFaceToFace + 2 * parInsulationThickness, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
    '  Insert your code for the bottom Hexagonal head 2
         oStPoint.Set 0, flangeDiam / 2 + 0.1 * parTrapHeight, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parFaceToFace * Tan(PI / 6) + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parFaceToFace * Tan(PI / 6) + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVect.Set 0, 1, 0
        Dim ObjHexagonalHead1 As Object
        Set ObjHexagonalHead1 = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parTrapHeight - parOffset - flangeDiam / 2 - 0.1 * parFaceToFace + parInsulationThickness, True)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexagonalHead1
        Set ObjHexagonalHead1 = Nothing
        Set oLineString = Nothing
        
    ElseIf lPartDataBasis = 428 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
    ' Insert your code for Output 1 (Cylinder)
        oStPoint.Set -parFaceToFace / 2, 0, 0
        oEnPoint.Set parFaceToFace / 2, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * parOffset1 + 2 * parInsulationThickness, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
    ' Insert your code for the Output 2 (Trap Box)
        oStPoint.Set -parFaceToFace / 2 + flangeThick - parInsulationThickness, 0.6 * flangeDiam + parInsulationThickness, 0.6 * flangeDiam + parInsulationThickness
        oEnPoint.Set parFaceToFace / 2 - flangeThick + parInsulationThickness, -0.6 * flangeDiam - parInsulationThickness, -0.6 * flangeDiam - parInsulationThickness
        Dim objTrapBox As Object
        Set objTrapBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTrapBox
        Set objTrapBox = Nothing
        
    '  Insert your code for the Top Hexagonal head
         dAngle1 = PI / 3
         If CmpDblEqual(parPipingSpecialtyWidth, 0) Then parPipingSpecialtyWidth = 0.5 * parFaceToFace
         oStPoint.Set 0, 0.6 * flangeDiam, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parPipingSpecialtyWidth / 2 + parInsulationThickness) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parPipingSpecialtyWidth / 2 + parInsulationThickness) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New IngrGeom3D.LineString3d
        Set oAxisVect = New AutoMath.DVector
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset2 - 0.6 * flangeDiam + parInsulationThickness, True)

    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
   
    End If

' Place Insulation for Port 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set parFaceToFace / 2, 0, 0
    oEnPoint.Set parFaceToFace / 2 - flangeThick - parInsulationThickness, 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Dim ObjInsulatedPort1 As Object
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

'   Set the output
   
    m_OutputColl.AddOutput "InsPort_", ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Place Insulation for Port 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    oStPoint.Set -parFaceToFace / 2, 0, 0
    oEnPoint.Set -parFaceToFace / 2 + flangeThick + parInsulationThickness, 0, 0
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Dim ObjInsulatedPort2 As Object
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

'   Set the output
   
    m_OutputColl.AddOutput "InsPort_", ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
