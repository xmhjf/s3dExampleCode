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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DOPSpurGear.OPSpurGear
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday, Sept 10 2007
'   Description     :  Spur Gear type operator to be used with 3 way diverter combination valve
'                      of Tyco Flow Control
'   Reference       :  http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   10.Sep.2007     PK          CR-126718:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parOperatorHeight       As Double
    Dim parArmLength            As Double
    Dim parArmLength1           As Double
    Dim parPinionShaftDiameter  As Double
    
    Dim iOutput              As Integer
    Dim dFlangeThickness     As Double
    Dim dStemVisibleHeight   As Double
    Dim dStemJamNutThickness As Double
    Dim dYokeHeight          As Double
    Dim dFlangeCenterHeight  As Double
    Dim dFlangeWidth         As Double
        
    Dim oCentPos   As AutoMath.DPosition
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    
    Set oCentPos = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
             
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parArmLength = arrayOfInputs(3)
    parArmLength1 = arrayOfInputs(4)
    parPinionShaftDiameter = arrayOfInputs(5)
    
    iOutput = 0
    oCentPos.Set 0, 0, 0
      
    'Insert your code for Flange(Output 1)
    Dim oFlange As Object
    
    'Assumptions:1)Connection point of two Flanges is assumed to be at 25% of Operator Height
    '            2)Flange Width is assumed to be 40% of the sum of Arm Length and Arm Length1
    '            3)Flange Thickness is assumed to be 1/32 times of Operator Height
    dFlangeCenterHeight = 0.25 * parOperatorHeight
    dFlangeWidth = 0.4 * (parArmLength + parArmLength1)
    dFlangeThickness = parOperatorHeight / 32
    
    oStPoint.Set -dFlangeWidth / 2 + oCentPos.x, oCentPos.y + dFlangeCenterHeight - dFlangeThickness, oCentPos.z + dFlangeWidth / 2
    oEnPoint.Set dFlangeWidth / 2 + oCentPos.x, oCentPos.y + dFlangeCenterHeight + dFlangeThickness, oCentPos.z - dFlangeWidth / 2
    Set oFlange = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       
    'Set Output 1
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oFlange
    Set oFlange = Nothing
    
    'Insert your code for Body(Output 2)
    Dim oBody As Object
    
    'Assumptions:1)Height of Body is assumed to be 70% of Operator Height
    dYokeHeight = 0.7 * parOperatorHeight
    oStPoint.Set oCentPos.x, oCentPos.y, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + dYokeHeight, oCentPos.z
    Set oBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.7 * dFlangeWidth, True)
        
    'Set Output  2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oBody = Nothing
    
    'Insert your code for Gear Casing(Output 3)
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oCStr           As IngrGeom3D.ComplexString3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oAxis           As New AutoMath.DVector
    Dim oProjection     As Object
    Dim opoints(3)      As New AutoMath.DPosition
    Dim iCount          As Integer
       
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    opoints(0).Set oCentPos.x + dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1
    opoints(1).Set oCentPos.x + dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z - 0.8 * parArmLength1
    opoints(2).Set oCentPos.x + parArmLength1, oCentPos.y + dYokeHeight, oCentPos.z
    
    Set oArc = PlaceTrArcBy3Pts(opoints(0), opoints(1), opoints(2))
    lines.Add oArc
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x + dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z - 0.8 * parArmLength1, oCentPos.x - dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z - 0.8 * parArmLength1)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z - 0.8 * parArmLength1, oCentPos.x - parArmLength, oCentPos.y + dYokeHeight, oCentPos.z - 0.8 * parArmLength1 + Tan(PI / 9) * parArmLength)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - parArmLength, oCentPos.y + dYokeHeight, oCentPos.z - 0.8 * parArmLength1 + Tan(PI / 9) * parArmLength, oCentPos.x - parArmLength, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1 - Tan(PI / 9) * parArmLength)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - parArmLength, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1 - Tan(PI / 9) * parArmLength, oCentPos.x - dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1, oCentPos.x + dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1)
    lines.Add oLine
    
    oStPoint.Set oCentPos.x + dFlangeWidth / 5, oCentPos.y + dYokeHeight, oCentPos.z + 0.8 * parArmLength1
    Set oCStr = PlaceTrCString(oStPoint, lines)
      
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
     
    oAxis.Set 0, 1, 0
    Set oProjection = PlaceProjection(m_OutputColl, oCStr, oAxis, 0.1 * parOperatorHeight, True)
    
    'Set Output 3
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oProjection
    
    Set oAxis = Nothing
    For iCount = 0 To 2
        Set opoints(iCount) = Nothing
    Next iCount
    Set oLine = Nothing
    Set oArc = Nothing
    Set lines = Nothing
    Set oCStr = Nothing
    Set oProjection = Nothing
    Set oGeomFactory = Nothing
   
    'Insert your code for Visible Stem(Output 4)
    Dim oVisibleStem As Object
    
    'Assumptions:1)Stem Visible Height is assumed to be 20% of Operator Height
    dStemVisibleHeight = 0.2 * parOperatorHeight
    oStPoint.Set oCentPos.x, oCentPos.y + dYokeHeight + 0.1 * parOperatorHeight, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + dYokeHeight + 0.1 * parOperatorHeight + dStemVisibleHeight, oCentPos.z
    Set oVisibleStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dFlangeWidth / 5, True)
    
    'Set Output 4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVisibleStem
    Set oVisibleStem = Nothing
    
    'Insert your code for Stem Jam Nut(Output5)
    Dim oStemJamNut As Object
        
    'Assumptions:1)Stem Jam Nut Thickness is assumed to be 1/20 times  of Operator Height
    dStemJamNutThickness = 0.05 * parOperatorHeight
    oStPoint.Set oCentPos.x, oCentPos.y + dYokeHeight + 0.2 * parOperatorHeight, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + dYokeHeight + 0.2 * parOperatorHeight + dStemJamNutThickness, oCentPos.z
    Set oStemJamNut = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 4 * dFlangeWidth / 5, True)
        
    'Set Output 5
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oStemJamNut
    Set oStemJamNut = Nothing
    
    'Insert your code for Pinion Shaft(Output 6)
    Dim oPinionShaft As Object
    oStPoint.Set oCentPos.x - parArmLength + 1.5 * parPinionShaftDiameter, oCentPos.y + dYokeHeight - 0.05 * parOperatorHeight, oCentPos.z
    oEnPoint.Set oCentPos.x - parArmLength + 1.5 * parPinionShaftDiameter, oCentPos.y + dYokeHeight + 0.15 * parOperatorHeight, oCentPos.z
    Set oPinionShaft = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPinionShaftDiameter, True)
    
    'Set Output 6
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oPinionShaft
    Set oPinionShaft = Nothing
    
    
    Set oCentPos = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
      
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

