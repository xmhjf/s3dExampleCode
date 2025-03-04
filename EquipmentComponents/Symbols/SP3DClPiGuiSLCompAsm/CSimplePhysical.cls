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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:45 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, May 13, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This is a Clamped Trunnion Pipe Guide for small Lines through 6 inch pipe.
'   This symbol consists of 4 inputs and 21 outputs,It has two aspects one is Simple Physical
'   another is ReferenceGeometry.There are 15 Physical and 5 Reference Outputs..
'   This is a Dow symbol Equipment
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double

Dim m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

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
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parVesselDiameter As Double
    Dim parPipeCLtoTowerCL As Double
    Dim parTowertoGuideBase As Double
    Dim parPipeOD As Double
        
    Dim iOutput     As Double
    
    Dim ObjSaddle1 As Object
    Dim ObjSaddle2 As Object
    Dim ObjSaddleSupport1 As Object
    Dim ObjSaddleSupport2 As Object
    Dim ObjSaddleSupport3 As Object
    Dim ObjSupportCenterPlate1 As Object
    Dim ObjGuideSupport1 As Object
    Dim ObjGuideSupport2 As Object
    Dim ObjGuideSupport3 As Object
    Dim ObjSupportCenterPlate2 As Object
    Dim ObjZClamp1A As Object
    Dim ObjZClamp1B As Object
    Dim ObjZClamp2A As Object
    Dim ObjZClamp2B As Object
'    Dim ObjBolt1 As Object
'    Dim ObjBolt2 As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPipeCLtoTowerCL = arrayOfInputs(3)
    parTowertoGuideBase = arrayOfInputs(4)
    parPipeOD = arrayOfInputs(5)
    
'Assumptions( these are made as per the sketch.)

''The part which guide the pipe is termed as Guide and The part which fixes to Tower
''is termed as TowerSuport in the naming convention

    ''Assumed SaddleThickness as 3/16th inch
    ''Assumed Support Plate thickness as 3/8inch.
    ''Assumed trapezoidal SupprotPlate of saddle angle as 90 deg
    ''Assumed support plate Height as 8inch (from sketch)
    ''Assumed Saddle Height as 12 inch( from sketch)
   
    m_oSymGeomHelper.OutputCollection = m_OutputColl

 ' Insert your code for output 1(Saddle1 Body)

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oEleCollection           As Collection
    Set oEleCollection = New Collection
    
    Dim dSaddleThickness As Double
    Dim dSaddleClampLength As Double
    Dim dSaddleHeight As Double
    Dim dGap As Double
    Dim dBoltBarThickness As Double
    dSaddleThickness = 0.005 ''3/16th inch
    dSaddleHeight = 0.304  ''12 Inch
    dSaddleClampLength = 0.057       '2-1/4 inch
    dGap = 0.003  '1/8 inch
    dBoltBarThickness = 0.013 '1/2 inch
    
    Dim dlinepts() As Double
    ReDim dlinepts(0 To 32) As Double
    
    ''Saddle Points
    '1
    dlinepts(0) = -parPipeOD / 2
    dlinepts(1) = parPipeCLtoTowerCL - dGap / 2
    dlinepts(2) = dSaddleHeight / 2
    
    '2
    dlinepts(3) = -parPipeOD / 2 - dSaddleClampLength
    dlinepts(4) = dlinepts(1)
    dlinepts(5) = dlinepts(2)
    
    '3
    dlinepts(6) = dlinepts(3)
    dlinepts(7) = parPipeCLtoTowerCL - dGap / 2 - dBoltBarThickness
    dlinepts(8) = dlinepts(2)
    
    '4
    dlinepts(9) = -parPipeOD / 2 - dSaddleThickness
    dlinepts(10) = dlinepts(7)
    dlinepts(11) = dlinepts(2)
    
    '5
    dlinepts(12) = 0
    dlinepts(13) = parPipeCLtoTowerCL - parPipeOD / 2 - dSaddleThickness
    dlinepts(14) = dlinepts(2)
    
    '6
    dlinepts(15) = parPipeOD / 2 + dSaddleThickness
    dlinepts(16) = dlinepts(7)
    dlinepts(17) = dlinepts(2)
    
    '7
    dlinepts(18) = parPipeOD / 2 + dSaddleClampLength
    dlinepts(19) = dlinepts(7)
    dlinepts(20) = dlinepts(2)
    
    '8
    dlinepts(21) = dlinepts(18)
    dlinepts(22) = dlinepts(1)
    dlinepts(23) = dlinepts(2)
    
    '9
    dlinepts(24) = parPipeOD / 2
    dlinepts(25) = dlinepts(1)
    dlinepts(26) = dlinepts(2)
    
    '10
    dlinepts(27) = 0
    dlinepts(28) = parPipeCLtoTowerCL - parPipeOD / 2
    dlinepts(29) = dlinepts(2)
    
    '11
    dlinepts(30) = dlinepts(0)
    dlinepts(31) = dlinepts(1)
    dlinepts(32) = dlinepts(2)
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(15), dlinepts(16), dlinepts(17))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                      dlinepts(18), dlinepts(19), dlinepts(20))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(18), dlinepts(19), dlinepts(20), _
                                                      dlinepts(21), dlinepts(22), dlinepts(23))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(21), dlinepts(22), dlinepts(23), _
                                                      dlinepts(24), dlinepts(25), dlinepts(26))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(24), dlinepts(25), dlinepts(26), _
                                                            dlinepts(27), dlinepts(28), dlinepts(29), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector
    
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1
    
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    
    Set ObjSaddle1 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dSaddleHeight, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddle1
    Set ObjSaddle1 = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    
    Dim count As Integer
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
 ' Insert your code for output 2(Saddle2 Body)
    
    ''Saddle Points
    '1
    dlinepts(0) = -parPipeOD / 2
    dlinepts(1) = parPipeCLtoTowerCL + dGap / 2
    dlinepts(2) = dSaddleHeight / 2

    '2
    dlinepts(3) = -parPipeOD / 2 - dSaddleClampLength
    dlinepts(4) = dlinepts(1)
    dlinepts(5) = dlinepts(2)
    
    '3
    dlinepts(6) = dlinepts(3)
    dlinepts(7) = parPipeCLtoTowerCL + dGap / 2 + dBoltBarThickness
    dlinepts(8) = dlinepts(2)
    
    '4
    dlinepts(9) = -parPipeOD / 2 - dSaddleThickness
    dlinepts(10) = dlinepts(7)
    dlinepts(11) = dlinepts(2)
    
    '5
    dlinepts(12) = 0
    dlinepts(13) = parPipeCLtoTowerCL + parPipeOD / 2 + dSaddleThickness
    dlinepts(14) = dlinepts(2)
    
    '6
    dlinepts(15) = parPipeOD / 2 + dSaddleThickness
    dlinepts(16) = dlinepts(7)
    dlinepts(17) = dlinepts(2)
    
    '7
    dlinepts(18) = parPipeOD / 2 + dSaddleClampLength
    dlinepts(19) = dlinepts(7)
    dlinepts(20) = dlinepts(2)
    
    '8
    dlinepts(21) = dlinepts(18)
    dlinepts(22) = dlinepts(1)
    dlinepts(23) = dlinepts(2)
    
    '9
    dlinepts(24) = parPipeOD / 2
    dlinepts(25) = dlinepts(1)
    dlinepts(26) = dlinepts(2)
    
    '10
    dlinepts(27) = 0
    dlinepts(28) = parPipeCLtoTowerCL + parPipeOD / 2
    dlinepts(29) = dlinepts(2)
    
    '11
    dlinepts(30) = dlinepts(0)
    dlinepts(31) = dlinepts(1)
    dlinepts(32) = dlinepts(2)
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(15), dlinepts(16), dlinepts(17))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                      dlinepts(18), dlinepts(19), dlinepts(20))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(18), dlinepts(19), dlinepts(20), _
                                                      dlinepts(21), dlinepts(22), dlinepts(23))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(21), dlinepts(22), dlinepts(23), _
                                                      dlinepts(24), dlinepts(25), dlinepts(26))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(24), dlinepts(25), dlinepts(26), _
                                                            dlinepts(27), dlinepts(28), dlinepts(29), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1
    
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    
    Set ObjSaddle2 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dSaddleHeight, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddle2
    Set ObjSaddle2 = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
'Insert your code for output 3(Top Trapizoidal Saddle Support)
    Dim dGuidePlateWidth As Double
    Dim dGuidePlateThick As Double
    Dim dGuidePlateHeight As Double
    Dim dGuidePlateThick1 As Double
     
    
    dGuidePlateThick = 0.006 '' 1/4 inch  (trapezoidal plate)
    dGuidePlateThick1 = 0.01     '3/8 inch
    dGuidePlateHeight = 0.204 ''8inch
    dGuidePlateWidth = 0.095 '' 3-3/4 inch

    ReDim dlinepts(0 To 14) As Double
    
    dlinepts(0) = Sin(PI / 4) * (parPipeOD / 2 + dSaddleThickness)
    dlinepts(1) = parPipeCLtoTowerCL - (Cos(PI / 4) * (parPipeOD / 2 + dSaddleThickness))
    dlinepts(2) = dGuidePlateHeight / 2

    dlinepts(3) = dGuidePlateWidth / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)

    dlinepts(6) = -dGuidePlateWidth / 2
    dlinepts(7) = dlinepts(4)
    dlinepts(8) = dlinepts(2)
    
    dlinepts(9) = -Sin(PI / 4) * (parPipeOD / 2 + dSaddleThickness)
    dlinepts(10) = dlinepts(1)
    dlinepts(11) = dlinepts(2)
     
    dlinepts(12) = 0
    dlinepts(13) = parPipeCLtoTowerCL - (parPipeOD / 2 + dSaddleThickness)
    dlinepts(14) = dlinepts(2)
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1
    
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
        
    Set ObjSaddleSupport1 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dGuidePlateThick, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddleSupport1
    Set ObjSaddleSupport1 = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
 ' Insert your code for output 4(Bottom Trapizoidal Saddle Support )
      
    dlinepts(0) = Sin(PI / 4) * (parPipeOD / 2 + dSaddleThickness)
    dlinepts(1) = parPipeCLtoTowerCL - (Cos(PI / 4) * (parPipeOD / 2 + dSaddleThickness))
    dlinepts(2) = -dGuidePlateHeight / 2

    dlinepts(3) = dGuidePlateWidth / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)
    
    dlinepts(6) = -dGuidePlateWidth / 2
    dlinepts(7) = dlinepts(4)
    dlinepts(8) = dlinepts(2)
    
    dlinepts(9) = -Sin(PI / 4) * (parPipeOD / 2 + dSaddleThickness)
    dlinepts(10) = parPipeCLtoTowerCL - (Cos(PI / 4) * (parPipeOD / 2 + dSaddleThickness))
    dlinepts(11) = dlinepts(2)
         
    dlinepts(12) = 0
    dlinepts(13) = parPipeCLtoTowerCL - (parPipeOD / 2 + dSaddleThickness)
    dlinepts(14) = dlinepts(2)
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    
    oAxisVec.Set 0, 0, 1
    Set ObjSaddleSupport2 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dGuidePlateThick, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddleSupport2
    Set ObjSaddleSupport2 = Nothing
        
    oComplexStr.RemoveCurve True
    
    Set oComplexStr = Nothing
    
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
' Insert your code for output 5(Box Type portion fits in Clamps)
    Dim oEnPoint As New AutoMath.DPosition
     ReDim dlinepts(0 To 14) As Double

'Point 1
    dlinepts(0) = dGuidePlateWidth / 2
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(2) = dGuidePlateHeight / 2
    'Point 2
    dlinepts(3) = dGuidePlateWidth / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = -dGuidePlateWidth / 2
    dlinepts(7) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = -dGuidePlateWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjSaddleSupport3 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dGuidePlateHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddleSupport3
    Set ObjSaddleSupport3 = Nothing
    
' Insert your code for output 6(Guide Center Plate)
    
    ReDim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dGuidePlateThick / 2
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(2) = dGuidePlateHeight / 2 - dGuidePlateThick
    'Point 2
    dlinepts(3) = dGuidePlateThick / 2
    dlinepts(4) = parPipeCLtoTowerCL - (parPipeOD / 2 + dSaddleThickness)
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = -dGuidePlateThick / 2
    dlinepts(7) = parPipeCLtoTowerCL - (parPipeOD / 2 + dSaddleThickness)
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = -dGuidePlateThick / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjSupportCenterPlate1 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dGuidePlateHeight - 2 * dGuidePlateThick, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupportCenterPlate1
    Set ObjSupportCenterPlate1 = Nothing
    
' Insert your code for output 7(Tower circular support)
    
    Dim dTowerSupportThick As Double
    Dim dTowerSaddleHeight As Double
    Dim dTowerSupportWidth As Double

    dTowerSupportThick = 0.006  ''1/4 inch  (Plate attached to Tower saddle in x-y plane)
    dTowerSaddleHeight = 0.254  ''10 inch  (As per note)
    dTowerSupportWidth = 0.216  ''8-1/2 inch

    '' Rib angle is calculated as per the dimensions given.
    Dim dAngle As Double ''Half of angle made by the tower top supprot plate at center of tower axis
    Dim dRibAngle As Double ''Half of angle made by the tower saddle at center of tower axis.
    dAngle = Atn((dTowerSupportWidth / 2) / (parVesselDiameter / 2))
    dRibAngle = 1.2 * dAngle ''Assumed 20% more of the rib angle for the circular portion that attached to tower.

    ReDim dlinepts(0 To 17) As Double

    ''Tower support Points
    dlinepts(0) = Sin(dRibAngle) * (parVesselDiameter / 2)
    dlinepts(1) = Cos(dRibAngle) * (parVesselDiameter / 2)
    dlinepts(2) = dTowerSaddleHeight / 3 'Asymetry assumed as 1:2

    dlinepts(3) = Sin(dRibAngle) * (parVesselDiameter / 2 + dTowerSupportThick)
    dlinepts(4) = Cos(dRibAngle) * (parVesselDiameter / 2 + dTowerSupportThick)
    dlinepts(5) = dlinepts(2)

    dlinepts(6) = 0
    dlinepts(7) = parVesselDiameter / 2 + dTowerSupportThick
    dlinepts(8) = dlinepts(2)

    dlinepts(9) = -Sin(dRibAngle) * (parVesselDiameter / 2 + dTowerSupportThick)
    dlinepts(10) = Cos(dRibAngle) * (parVesselDiameter / 2 + dTowerSupportThick)
    dlinepts(11) = dlinepts(2)

    dlinepts(12) = -Sin(dRibAngle) * (parVesselDiameter / 2)
    dlinepts(13) = Cos(dRibAngle) * (parVesselDiameter / 2)
    dlinepts(14) = dlinepts(2)

    dlinepts(15) = 0
    dlinepts(16) = parVesselDiameter / 2
    dlinepts(17) = dlinepts(2)

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                            dlinepts(6), dlinepts(7), dlinepts(8), _
                                                            dlinepts(9), dlinepts(10), dlinepts(11))
    oEleCollection.Add oArc
    Set oArc = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                      dlinepts(12), dlinepts(13), dlinepts(14))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(15), dlinepts(16), dlinepts(17), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)

    Set ObjGuideSupport1 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dTowerSaddleHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjGuideSupport1
    Set ObjGuideSupport1 = Nothing

    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing

    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count


'Insert your code for output 8(Top guide plate)
    Dim dClampPlateHeight As Double ''height of Plate on which clamps are mounted.
    dClampPlateHeight = 0.102 ''4 inch   ''Plate on which clamps are mounted.

    ReDim dlinepts(0 To 14) As Double

    dlinepts(0) = dTowerSupportWidth / 2
    dlinepts(1) = Cos(dAngle) * (parVesselDiameter / 2 + dTowerSupportThick)
    dlinepts(2) = dClampPlateHeight / 2

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase - dTowerSupportThick
    dlinepts(5) = dlinepts(2)

    dlinepts(6) = -dlinepts(0)
    dlinepts(7) = dlinepts(4)
    dlinepts(8) = dlinepts(2)

    dlinepts(9) = dlinepts(6)
    dlinepts(10) = Cos(dAngle) * (parVesselDiameter / 2 + dTowerSupportThick)
    dlinepts(11) = dlinepts(2)

    dlinepts(12) = 0
    dlinepts(13) = parVesselDiameter / 2 + dTowerSupportThick
    dlinepts(14) = dlinepts(2)

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))

    oEleCollection.Add oLine
    Set oLine = Nothing

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)

    Set ObjGuideSupport2 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dTowerSupportThick, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjGuideSupport2
    Set ObjGuideSupport2 = Nothing

    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing

    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

' Insert your code for output 9(Guide Clamp Plate)
   ReDim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dTowerSupportWidth / 2
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase - dTowerSupportThick
    dlinepts(2) = dClampPlateHeight / 2
    'Point 2
    dlinepts(3) = dTowerSupportWidth / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = -dTowerSupportWidth / 2
    dlinepts(7) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = -dTowerSupportWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase - dTowerSupportThick
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjGuideSupport3 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dClampPlateHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjGuideSupport3
    Set ObjGuideSupport3 = Nothing

' Insert your code for output 10(Guide support CenterPlate)
    ReDim dlinepts(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d

    dlinepts(0) = dTowerSupportThick / 2
    dlinepts(1) = parVesselDiameter / 2 + dTowerSupportThick
    dlinepts(2) = dClampPlateHeight / 2

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase - dTowerSupportThick
    dlinepts(5) = dlinepts(2)

    dlinepts(6) = dlinepts(0)
    dlinepts(7) = dlinepts(4)
    dlinepts(8) = -dlinepts(2)

    dlinepts(9) = dlinepts(0)
    dlinepts(10) = parVesselDiameter / 2 + dTowerSupportThick
    dlinepts(11) = -dClampPlateHeight

    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinepts)

    oAxisVec.Set -1, 0, 0
    Set ObjSupportCenterPlate2 = PlaceProjection(m_OutputColl, oLineString, oAxisVec, dTowerSupportThick, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupportCenterPlate2
    Set ObjSupportCenterPlate2 = Nothing

' Insert your code for output 11(Left Z-Clamp Inner Box)
''Assumed Clamp Length as 3 inch
''(10 % of inner portion of Z-clamp's Length is assumed extends out from the tower support plate.)
    Dim dZClampLength As Double
    
    dZClampLength = 0.0762 ''3 inch.
    ReDim dlinepts(0 To 14) As Double
    Dim dZClampExten As Double
    dZClampExten = 0.005

    'Point 1
    dlinepts(0) = dTowerSupportWidth / 2 + dZClampExten '(As per Skech)
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(2) = dClampPlateHeight / 2
    'Point 2
    dlinepts(3) = dTowerSupportWidth / 2 + dZClampExten  '(As per Skech)
    dlinepts(4) = dlinepts(1) + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = 1.1 * dGuidePlateWidth / 2
    dlinepts(7) = dlinepts(1) + dGuidePlateThick1
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = 1.1 * dGuidePlateWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjZClamp1A = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dClampPlateHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjZClamp1A
    Set ObjZClamp1A = Nothing
    
' Insert your code for output 12(Left Z-Clamp Outer Box)
  ReDim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dTowerSupportWidth / 2 - dZClampExten  '(As per Skech)
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(2) = 1.1 * (dClampPlateHeight / 2)
    'Point 2
    dlinepts(3) = dTowerSupportWidth / 2 - dZClampExten  '(As per Skech)
    dlinepts(4) = dlinepts(1) + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = dlinepts(3) - dZClampLength
    dlinepts(7) = dlinepts(1) + dGuidePlateThick1
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = dlinepts(3) - dZClampLength
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjZClamp1B = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, 1.1 * dClampPlateHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjZClamp1B
    Set ObjZClamp1B = Nothing
    
' Insert your code for output 13(Right Z-Clamp Inner Box)
  ReDim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = -dTowerSupportWidth / 2 - dZClampExten '(As per Skech)
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(2) = dClampPlateHeight / 2
    'Point 2
    dlinepts(3) = -dTowerSupportWidth / 2 - dZClampExten  '(As per Skech)
    dlinepts(4) = dlinepts(1) + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = -1.1 * dGuidePlateWidth / 2
    dlinepts(7) = dlinepts(1) + dGuidePlateThick1
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = -1.1 * dGuidePlateWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjZClamp2A = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dClampPlateHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjZClamp2A
    Set ObjZClamp2A = Nothing
    
' Insert your code for output 14(Right Z-Clamp outer Box)
   ReDim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = -dTowerSupportWidth / 2 + dZClampExten  '(As per Skech)
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(2) = 1.1 * (dClampPlateHeight / 2)
    'Point 2
    dlinepts(3) = -dTowerSupportWidth / 2 + dZClampExten  '(As per Skech)
    dlinepts(4) = dlinepts(1) + dGuidePlateThick1
    dlinepts(5) = dlinepts(2)
    'Point 3
    dlinepts(6) = dlinepts(3) + dZClampLength
    dlinepts(7) = dlinepts(1) + dGuidePlateThick1
    dlinepts(8) = dlinepts(2)
    'Point 4
    dlinepts(9) = dlinepts(3) + dZClampLength
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase + dGuidePlateThick1
    dlinepts(11) = dlinepts(2)
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    oAxisVec.Set 0, 0, -1
    Set ObjZClamp2B = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, 1.1 * dClampPlateHeight, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjZClamp2B
    Set ObjZClamp2B = Nothing
        
        
' Insert your code for output 15(Left Bolt)As a Cylinder
    Dim dBoltDia As Double
    Dim dBoltCenter As Double ''Distance beteween centers of bolts.
    Dim dBoltLength As Double
    
    dBoltDia = 0.018 ''3/4 inch
    dBoltCenter = 0.07 ''2 3/4 th inch.
    dBoltLength = 3 * dBoltDia + 2 * dGuidePlateThick + dTowerSupportThick
    
    oStPoint.Set dBoltCenter, _
                    parVesselDiameter / 2 + parTowertoGuideBase - (dTowerSupportThick + 2 * dBoltDia), 0
    oEnPoint.Set dBoltCenter, oStPoint.y + dBoltLength, 0
                      
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
     m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dBoltDia
    
' Insert your code for output 16(Right Bolt)As a Cylinder
    oStPoint.Set -dBoltCenter, _
                parVesselDiameter / 2 + parTowertoGuideBase - (dTowerSupportThick + 2 * dBoltDia), 0
                            
    oEnPoint.Set -dBoltCenter, oStPoint.y + dBoltLength, 0
                      
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dBoltDia
       
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oEleCollection = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                        dPoints() As Double, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean) As Object

    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler
      
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d
        
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = geomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLineStr, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    height, isCapped)
    
    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set geomFactory = Nothing
    
    Exit Function
ErrorHandler:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Function
