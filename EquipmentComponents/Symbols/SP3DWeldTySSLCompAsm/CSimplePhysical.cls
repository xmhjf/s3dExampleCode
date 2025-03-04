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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:38 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, 22 April 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol consists of 12 Physical outputs. Of Which 10 plates, 2 saddles
'  The symbol consists of 5 Reference Geometry output which are 1 Reference plane,2 lines and 2 points.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim ObjPipeSideSaddle As Object
    Dim ObjVesselSideSaddle As Object
    Dim ObjPlate1 As Object
    Dim ObjPlate2 As Object
    Dim ObjPlate3 As Object
    Dim ObjPlate4 As Object
    Dim ObjPlate5 As Object
    Dim ObjPlate6 As Object
    Dim ObjPlate7 As Object
    Dim ObjPlate8 As Object
    Dim ObjPlate9 As Object
    Dim ObjPlate10 As Object
    
    Dim parVesselDiameter As Double
    Dim parPipeCLtoTowerCL As Double
    Dim parTowertoGuideBase As Double
    Dim parPipeOD As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPipeCLtoTowerCL = arrayOfInputs(3)
    parTowertoGuideBase = arrayOfInputs(4)
    parPipeOD = arrayOfInputs(5)
    
     iOutput = 0
    
    ' The dow drawing indicates the saddle width on the pipe side as equal to Nominal Pipe Diameter with Max of 76mm(3 ")
    ' However Presently the Pipe side of the Saddle width is taken to be equal to the Pipe OD which is an input parameter(parPipeOD)
    Dim dPipeSideSaddleWidth As Double
    dPipeSideSaddleWidth = parPipeOD
    
    Dim dPipeSideSaddlehalfAngle As Double
    Dim dPipeSideSaddleThickness As Double
    Dim dPipeSideSaddleHeight As Double
    dPipeSideSaddleHeight = 0.356 ' 14 inch
    dPipeSideSaddleThickness = 0.006 ' 1/4 inch
    Dim saddleEdgetoCL As Double
    saddleEdgetoCL = Sqr(((parPipeOD / 2 + dPipeSideSaddleThickness) ^ 2) - ((dPipeSideSaddleWidth / 2) ^ 2))
    dPipeSideSaddlehalfAngle = Atn((dPipeSideSaddleWidth / 2) / saddleEdgetoCL)
    
' Insert your code for output 1(Pipe Side Saddle guide)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oEleCollection           As Collection
    Set oEleCollection = New Collection
    
    Dim dlinepts(0 To 17) As Double
            
    ''Saddle Points
    'Point 1
    dlinepts(0) = -(parPipeOD / 2 + dPipeSideSaddleThickness) * Sin(dPipeSideSaddlehalfAngle)
    dlinepts(1) = parPipeCLtoTowerCL - (parPipeOD / 2 + dPipeSideSaddleThickness) * Cos(dPipeSideSaddlehalfAngle)
    dlinepts(2) = dPipeSideSaddleHeight / 2
    'Point 2
    dlinepts(3) = -(parPipeOD / 2) * Sin(dPipeSideSaddlehalfAngle)
    dlinepts(4) = parPipeCLtoTowerCL - (parPipeOD / 2) * Cos(dPipeSideSaddlehalfAngle)
    dlinepts(5) = dPipeSideSaddleHeight / 2
    'Point 3
    dlinepts(6) = 0
    dlinepts(7) = parPipeCLtoTowerCL - parPipeOD / 2
    dlinepts(8) = dPipeSideSaddleHeight / 2
    'Point 4
    dlinepts(9) = (parPipeOD / 2) * Sin(dPipeSideSaddlehalfAngle)
    dlinepts(10) = parPipeCLtoTowerCL - (parPipeOD / 2) * Cos(dPipeSideSaddlehalfAngle)
    dlinepts(11) = dPipeSideSaddleHeight / 2
    'Point 5
    dlinepts(12) = (parPipeOD / 2 + dPipeSideSaddleThickness) * Sin(dPipeSideSaddlehalfAngle)
    dlinepts(13) = parPipeCLtoTowerCL - (parPipeOD / 2 + dPipeSideSaddleThickness) * Cos(dPipeSideSaddlehalfAngle)
    dlinepts(14) = dPipeSideSaddleHeight / 2
    'Point 6
    dlinepts(15) = 0
    dlinepts(16) = parPipeCLtoTowerCL - (parPipeOD / 2 + dPipeSideSaddleThickness)
    dlinepts(17) = dPipeSideSaddleHeight / 2
    
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
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1
    
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    Set ObjPipeSideSaddle = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPipeSideSaddleHeight, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeSideSaddle
    Set ObjPipeSideSaddle = Nothing
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    Dim count As Integer
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
' Insert your code for output 2(Plate 1)
    Dim dPlateBottWidth As Double
    Dim dPlateThk1 As Double
    Dim dPlateThk2 As Double
    Dim dGuidehalfHeight As Double
    Dim dGuideWidth As Double
    dPlateBottWidth = 0.132 ' 6"- 2*3/8" inch
    dPlateThk1 = 0.006 ' 1/4inch
    dPlateThk2 = 0.013 ' 1/2inch
    dGuidehalfHeight = 0.152 ' 6 inch
    dGuideWidth = 0.152 ' 6 inch
    
    'Point 1
    dlinepts(0) = -(parPipeOD / 2 + dPipeSideSaddleThickness) * Sin(dPipeSideSaddlehalfAngle)
    dlinepts(1) = parPipeCLtoTowerCL - (parPipeOD / 2 + dPipeSideSaddleThickness) * Cos(dPipeSideSaddlehalfAngle)
    dlinepts(2) = dGuidehalfHeight
    'Point 2
    dlinepts(3) = 0
    dlinepts(4) = parPipeCLtoTowerCL - (parPipeOD / 2 + dPipeSideSaddleThickness)
    dlinepts(5) = dGuidehalfHeight
    'Point 3
    dlinepts(6) = (parPipeOD / 2 + dPipeSideSaddleThickness) * Sin(dPipeSideSaddlehalfAngle)
    dlinepts(7) = parPipeCLtoTowerCL - (parPipeOD / 2 + dPipeSideSaddleThickness) * Cos(dPipeSideSaddlehalfAngle)
    dlinepts(8) = dGuidehalfHeight
    'Point 4
    dlinepts(9) = dPlateBottWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase + dPlateThk2
    dlinepts(11) = dGuidehalfHeight
    'Point 5
    dlinepts(12) = -dPlateBottWidth / 2
    dlinepts(13) = parVesselDiameter / 2 + parTowertoGuideBase + dPlateThk2
    dlinepts(14) = dGuidehalfHeight

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                            dlinepts(3), dlinepts(4), dlinepts(5), _
                                                            dlinepts(6), dlinepts(7), dlinepts(8))
    oEleCollection.Add oArc
    Set oArc = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    oEleCollection.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                      dlinepts(12), dlinepts(13), dlinepts(14))
    oEleCollection.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                      dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oLine
    Set oLine = Nothing
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
        
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    
    Set ObjPlate1 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPlateThk1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate1
    Set ObjPlate1 = Nothing
    
' Insert your code for output 3(Plate 2)
    ' The complex string created for the top plate(plate 1) is translated down to use it for projection for the middle and bottom plates.
    Dim translateMat     As New AutoMath.DT4x4   'For complexstring translation
    Dim zTransVec As New AutoMath.DVector
    zTransVec.Set 0, 0, -1
    zTransVec.Length = dGuidehalfHeight - dPlateThk1 / 2
    translateMat.LoadIdentity
    translateMat.Translate zTransVec
    
    'Transform the complex string to the new location
    oComplexStr.Transform translateMat
    Set ObjPlate2 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPlateThk1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate2
    Set ObjPlate2 = Nothing

' Insert your code for output 4(Plate 3)
    zTransVec.Length = dGuidehalfHeight - dPlateThk1 / 2
    translateMat.LoadIdentity
    translateMat.Translate zTransVec
    
    'Transform the complex string to the new location
    oComplexStr.Transform translateMat
    Set ObjPlate3 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPlateThk1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate3
    Set ObjPlate3 = Nothing

    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

' Insert your code for output 5(Plate 4)Bolted Plate on Pipe side
    'Point 1
    dlinepts(0) = -dGuideWidth / 2
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(2) = dGuidehalfHeight
    'Point 2
    dlinepts(3) = -dGuideWidth / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase + dPlateThk2
    dlinepts(5) = dGuidehalfHeight
    'Point 3
    dlinepts(6) = dGuideWidth / 2
    dlinepts(7) = parVesselDiameter / 2 + parTowertoGuideBase + dPlateThk2
    dlinepts(8) = dGuidehalfHeight
    'Point 4
    dlinepts(9) = dGuideWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(11) = dGuidehalfHeight
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Set ObjPlate4 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, 2 * dGuidehalfHeight, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate4
    Set ObjPlate4 = Nothing

' Insert your code for output 6(Plate 5)Vertical Mid plate
'Point 1
    dlinepts(0) = -dPlateThk1 / 2
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase + dPlateThk2
    dlinepts(2) = dGuidehalfHeight - dPlateThk1
    'Point 2
    dlinepts(3) = -dPlateThk1 / 2
    dlinepts(4) = parPipeCLtoTowerCL - parPipeOD / 2 - dPipeSideSaddleThickness
    dlinepts(5) = dGuidehalfHeight - dPlateThk1
    'Point 3
    dlinepts(6) = dPlateThk1 / 2
    dlinepts(7) = parPipeCLtoTowerCL - parPipeOD / 2 - dPipeSideSaddleThickness
    dlinepts(8) = dGuidehalfHeight - dPlateThk1
    'Point 4
    dlinepts(9) = dPlateThk1 / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase + dPlateThk2
    dlinepts(11) = dGuidehalfHeight - dPlateThk1
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Set ObjPlate5 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, 2 * dGuidehalfHeight - 2 * dPlateThk1, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate5
    Set ObjPlate5 = Nothing

' Insert your code for output 7(Vessel side guide)
    Dim dVessSideSaddleWidth As Double
    Dim dVessSideSaddlehalfAngle As Double
    Dim dVessSideSaddleThickness As Double
    Dim dVessSideSaddleHeight As Double
    dVessSideSaddleHeight = 0.356 ' 14 inch
    dVessSideSaddleThickness = 0.006 '1/4 inch
    dVessSideSaddleWidth = 0.132 '6"- 2*3/8" inch
    'Determine the angle of the Vessel Side guide saddle.
    Dim dPlatehalfAngle As Double ' This is the half angle which the plate makes in contact with the saddle on the vessel side
    dPlatehalfAngle = Atn((dVessSideSaddleWidth / 2) / (parVesselDiameter / 2 + dVessSideSaddleThickness))
    'Taking the extra saddle portion to be 1.2 times more the plate half angle
    dVessSideSaddlehalfAngle = dPlatehalfAngle * 1.2
    
    'Saddle Points
    'Point 1
    dlinepts(0) = -(parVesselDiameter / 2 + dVessSideSaddleThickness) * Sin(dVessSideSaddlehalfAngle)
    dlinepts(1) = (parVesselDiameter / 2 + dVessSideSaddleThickness) * Cos(dVessSideSaddlehalfAngle)
    dlinepts(2) = dVessSideSaddleHeight / 2
    'Point 2
    dlinepts(3) = -(parVesselDiameter / 2) * Sin(dVessSideSaddlehalfAngle)
    dlinepts(4) = (parVesselDiameter / 2) * Cos(dVessSideSaddlehalfAngle)
    dlinepts(5) = dVessSideSaddleHeight / 2
    'Point 3
    dlinepts(6) = 0
    dlinepts(7) = parVesselDiameter / 2
    dlinepts(8) = dVessSideSaddleHeight / 2
    'Point 4
    dlinepts(9) = (parVesselDiameter / 2) * Sin(dVessSideSaddlehalfAngle)
    dlinepts(10) = (parVesselDiameter / 2) * Cos(dVessSideSaddlehalfAngle)
    dlinepts(11) = dVessSideSaddleHeight / 2
    'Point 5
    dlinepts(12) = (parVesselDiameter / 2 + dVessSideSaddleThickness) * Sin(dVessSideSaddlehalfAngle)
    dlinepts(13) = (parVesselDiameter / 2 + dVessSideSaddleThickness) * Cos(dVessSideSaddlehalfAngle)
    dlinepts(14) = dVessSideSaddleHeight / 2
    'Point 6
    dlinepts(15) = 0
    dlinepts(16) = (parVesselDiameter / 2 + dVessSideSaddleThickness)
    dlinepts(17) = dVessSideSaddleHeight / 2

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
    
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    Set ObjVesselSideSaddle = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dVessSideSaddleHeight, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselSideSaddle
    Set ObjVesselSideSaddle = Nothing

    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

' Insert your code for output 8(Plate 6)
    'Point 1
    dlinepts(0) = -(parVesselDiameter / 2 + dVessSideSaddleThickness) * Sin(dPlatehalfAngle)
    dlinepts(1) = (parVesselDiameter / 2 + dVessSideSaddleThickness) * Cos(dPlatehalfAngle)
    dlinepts(2) = dGuidehalfHeight
    'Point 2
    dlinepts(3) = 0
    dlinepts(4) = (parVesselDiameter / 2 + dVessSideSaddleThickness)
    dlinepts(5) = dGuidehalfHeight
    'Point 3
    dlinepts(6) = (parVesselDiameter / 2 + dVessSideSaddleThickness) * Sin(dPlatehalfAngle)
    dlinepts(7) = (parVesselDiameter / 2 + dVessSideSaddleThickness) * Cos(dPlatehalfAngle)
    dlinepts(8) = dGuidehalfHeight
    'Point 4
    dlinepts(9) = dPlateBottWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase - dPlateThk2
    dlinepts(11) = dGuidehalfHeight
    'Point 5
    dlinepts(12) = -dPlateBottWidth / 2
    dlinepts(13) = parVesselDiameter / 2 + parTowertoGuideBase - dPlateThk2
    dlinepts(14) = dGuidehalfHeight

    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                            dlinepts(3), dlinepts(4), dlinepts(5), _
                                                            dlinepts(6), dlinepts(7), dlinepts(8))
    oEleCollection.Add oArc
    Set oArc = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    oEleCollection.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                      dlinepts(12), dlinepts(13), dlinepts(14))
    oEleCollection.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                      dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oLine
    Set oLine = Nothing
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set 0, 0, -1

    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    Set ObjPlate6 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPlateThk1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate6
    Set ObjPlate6 = Nothing
    Set oStPoint = Nothing

' Insert your code for output 9(Plate 7)
   ' The complex string created for the top plate(plate 1) is translated down to use it for projection for the middle and bottom plates.
    zTransVec.Length = dGuidehalfHeight - dPlateThk1 / 2
    translateMat.LoadIdentity
    translateMat.Translate zTransVec
    
    'Transform the complex string to the new location
    oComplexStr.Transform translateMat
    Set ObjPlate7 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPlateThk1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate7
    Set ObjPlate7 = Nothing

' Insert your code for output 10(Plate 8)
    zTransVec.Length = dGuidehalfHeight - dPlateThk1 / 2
    translateMat.LoadIdentity
    translateMat.Translate zTransVec
    
    'Transform the complex string to the new location
    oComplexStr.Transform translateMat
    Set ObjPlate8 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dPlateThk1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate8
    Set ObjPlate8 = Nothing
    Set translateMat = Nothing
    Set zTransVec = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count

' Insert your code for output 11(Plate 9)Bolted Plate on Vessel side
    'Point 1
    dlinepts(0) = -dGuideWidth / 2
    dlinepts(1) = parVesselDiameter / 2 + parTowertoGuideBase - dPlateThk2
    dlinepts(2) = dGuidehalfHeight
    'Point 2
    dlinepts(3) = -dGuideWidth / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(5) = dGuidehalfHeight
    'Point 3
    dlinepts(6) = dGuideWidth / 2
    dlinepts(7) = parVesselDiameter / 2 + parTowertoGuideBase
    dlinepts(8) = dGuidehalfHeight
    'Point 4
    dlinepts(9) = dGuideWidth / 2
    dlinepts(10) = parVesselDiameter / 2 + parTowertoGuideBase - dPlateThk2
    dlinepts(11) = dGuidehalfHeight
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Set ObjPlate9 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, 2 * dGuidehalfHeight, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate9
    Set ObjPlate9 = Nothing

' Insert your code for output 12(Plate 10)Vertical Mid plate
    'Point 1
    dlinepts(0) = -dPlateThk1 / 2
    dlinepts(1) = parVesselDiameter / 2 + dVessSideSaddleThickness
    dlinepts(2) = dGuidehalfHeight - dPlateThk1
    'Point 2
    dlinepts(3) = -dPlateThk1 / 2
    dlinepts(4) = parVesselDiameter / 2 + parTowertoGuideBase - dPlateThk2
    dlinepts(5) = dGuidehalfHeight - dPlateThk1
    'Point 3
    dlinepts(6) = dPlateThk1 / 2
    dlinepts(7) = parVesselDiameter / 2 + parTowertoGuideBase - dPlateThk2
    dlinepts(8) = dGuidehalfHeight - dPlateThk1
    'Point 4
    dlinepts(9) = dPlateThk1 / 2
    dlinepts(10) = parVesselDiameter / 2 + dVessSideSaddleThickness
    dlinepts(11) = dGuidehalfHeight - dPlateThk1
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Set ObjPlate10 = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, 2 * dGuidehalfHeight - 2 * dPlateThk1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate10
    Set ObjPlate10 = Nothing
    Set oAxisVec = Nothing

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



