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
'  Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         KKC
'   Creation Date:  Tuesday, Sep 18 2007
'   Source:
'    For partdatabasis 425:Thermodynamic Steam Traps Product Brochure, Model TD32F; Copyright 2004 Spirax Sarco®, Inc.
'       www.spiraxsarco.com/us
'    For partdatabasis 426:Thermodynamic Steam Traps Product Brochure, Model TD3-3; Copyright 2004 Spirax Sarco®, Inc.
'       www.spiraxsarco.com/us
'    For partdatabasis 427:Thermodynamic Steam Traps Product Brochure, Model TD259; Copyright 2004 Spirax Sarco®, Inc.
'        www.spiraxsarco.com/us
'    For partdatabasis 428:Thermodynamic Steam Traps Product Brochure, Model TD52; Copyright 2004 Spirax Sarco®, Inc.
'         www.spiraxsarco.com/us
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------     -----    ------------------
'   17.Dec.2007     KKC      CR-121408  Create additional symbols required by Intergraph's DIN piping specs
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
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
        dProjLength = 0.6 * (parFaceToFace - 2 * flangeThick)
        dAngle = Atn(parOffset2 / dProjLength)
        
        If CmpDblEqual(parDiameter1, 0) Then parDiameter1 = 0.25 * parFaceToFace
        If CmpDblEqual(parDiameter2, 0) Then parDiameter2 = parDiameter1
    ' Insert your code for Output 1 (Cylinder)
        oStPoint.Set -parFaceToFace / 2, 0, 0
        oEnPoint.Set parFaceToFace / 2, 0, 0
        Dim objCylinder As Object
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, 0, 0
        
        Dim objCylinder1 As Object
        Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * pipeDiam, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder1
        Set objCylinder1 = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
       
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, -(parOffset2 - (parDiameter2 / 2) * Cos(dAngle)), 0
        
        Dim objInclCylinder As Object
        Set objInclCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter2, True)
        
    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", objInclCylinder
        Set objInclCylinder = Nothing
    ' Insert your code for the Output 4 (Hexagonal Head)
         Dim dLineStrPoints(0 To 20)  As Double
         Dim oLineString As IngrGeom3D.LineString3d
         Set oLineString = New IngrGeom3D.LineString3d
         Dim dAngle1 As Double
         dAngle1 = PI / 3
         oStPoint.Set 0, 0.75 * pipeDiam, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parDiameter1 / 2) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parDiameter1 / 2) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        Dim oAxisVect As AutoMath.DVector
        Dim ObjHexagonalHead As Object
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset1 - 0.75 * pipeDiam, True)

    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oAxisVect = Nothing
        
    ElseIf lPartDataBasis = 426 Then
    
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        dProjLength = 0.5 * (parFaceToFace - 2 * flangeThick)
        dAngle = Atn(parOffset / dProjLength)
        
        If CmpDblEqual(parDiameter1, 0) Then parDiameter1 = 0.25 * parFaceToFace
        If CmpDblEqual(parDiameter2, 0) Then parDiameter2 = 0.4 * parDiameter1
        
    ' Insert your code for Output 1 (Cylinder)
        oStPoint.Set -parFaceToFace / 2, 0, 0
        oEnPoint.Set parFaceToFace / 2, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
    ' Insert your code for the Output 2 (Cylinder 2)
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, 0, 0
        
        Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam, True)
    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", objCylinder1
        Set objCylinder1 = Nothing
        
    ' Insert your code for the Output 3 (Inclined Cylinder)
       
        oStPoint.Set -dProjLength / 2, 0, 0
        oEnPoint.Set dProjLength / 2, -(parOffset2 - (parDiameter2 / 2 * Cos(dAngle))), 0
        
        Set objInclCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter2, True)
        
    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", objInclCylinder
        Set objInclCylinder = Nothing
        
    ' Insert your code for the Output 4 ( Hexagonal Head)
         dAngle1 = PI / 3
         oStPoint.Set 0, 0.6 * pipeDiam, 0
'       First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parDiameter1 / 2) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parDiameter1 / 2) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset1 - 0.6 * pipeDiam, True)

    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
        Set oLineString = Nothing
        Dim ObjHexagonalNut As Object
        Dim ObjHexagonalNut1 As Object
        
        oStPoint.Set -dProjLength / 2 - 0.1 * parFaceToFace, 0, 0
    '   First point is considered on positive Z-axis at an angle (PI / 6 + Angle) from positive Y-axis.
        iCount = 0
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x
            dLineStrPoints(3 * iCount - 2) = oStPoint.y + (0.8 * pipeDiam) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (0.8 * pipeDiam) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oLineString = New IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 1, 0, 0
        
        Set ObjHexagonalNut = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.1 * parFaceToFace, True)
        Set ObjHexagonalNut1 = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.1 * parFaceToFace, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexagonalNut
        Set ObjHexagonalNut = Nothing
        
        Dim oTransMatrix As AutoMath.DT4x4
        Set oTransMatrix = New AutoMath.DT4x4
        
        oTransMatrix.LoadIdentity
        Dim oDir As AutoMath.DVector
        Set oDir = New AutoMath.DVector
        
        oDir.Set dProjLength + 0.1 * parFaceToFace, 0, 0
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
         oStPoint.Set 0, -parOffset, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parFaceToFace * Tan(PI / 6)) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parFaceToFace * Tan(PI / 6)) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New IngrGeom3D.LineString3d
        Set oAxisVect = New AutoMath.DVector
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVect.Set 0, 1, 0
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset + flangeDiam / 2, True)

    '   Set the output
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
        
    '  Inset your code for the Output 2 (Cylinder)
        oStPoint.Set 0, flangeDiam / 2, 0
        oEnPoint.Set 0, flangeDiam / 2 + 0.1 * parTrapHeight, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.7 * parFaceToFace, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
    '  Insert your code for the bottom Hexagonal head 2
         oStPoint.Set 0, flangeDiam / 2 + 0.1 * parTrapHeight, 0
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parFaceToFace * Tan(PI / 6)) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parFaceToFace * Tan(PI / 6)) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVect.Set 0, 1, 0
        Dim ObjHexagonalHead1 As Object
        Set ObjHexagonalHead1 = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parTrapHeight - parOffset - flangeDiam / 2 - 0.1 * parFaceToFace, True)

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
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * parOffset1, True)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
    ' Insert your code for the Output 2 (Trap Box)
        oStPoint.Set -parFaceToFace / 2 + flangeThick, 0.6 * flangeDiam, 0.6 * flangeDiam
        oEnPoint.Set parFaceToFace / 2 - flangeThick, -0.6 * flangeDiam, -0.6 * flangeDiam
        Dim objTrapBox As Object
        Set objTrapBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTrapBox
        Set objTrapBox = Nothing
        
    '  Insert your code for the Top Hexagonal head
         oStPoint.Set 0, 0.6 * flangeDiam, 0
         dAngle1 = PI / 3
         If CmpDblEqual(parPipingSpecialtyWidth, 0) Then parPipingSpecialtyWidth = 0.5 * parFaceToFace
    '   First point is considered on positive X-axis at an angle (PI / 6 + Angle) from positive Z-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x + (parPipingSpecialtyWidth / 2) * Cos(PI / 6 + iCount * dAngle1)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y
            dLineStrPoints(3 * iCount - 1) = oStPoint.z + (parPipingSpecialtyWidth / 2) * Sin(PI / 6 + iCount * dAngle1)
        Next iCount
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New IngrGeom3D.LineString3d
        Set oAxisVect = New AutoMath.DVector
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        oAxisVect.Set 0, 1, 0
        Set ObjHexagonalHead = PlaceProjection(m_OutputColl, oLineString, oAxisVect, parOffset2 - 0.6 * flangeDiam, True)

    '   Set the output
      
        m_OutputColl.AddOutput "TrapBody_", ObjHexagonalHead
        Set ObjHexagonalHead = Nothing
   
    End If

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFaceToFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set parFaceToFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
