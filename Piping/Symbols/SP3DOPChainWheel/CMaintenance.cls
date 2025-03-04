VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 1 2004
'   Description:
''  Operator Chain Wheel symbol is prepared based on SP3D Concept design document
''  "Valve Accessories"(valveaccessoriescd.doc).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim PI As Double

Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()

    PI = Atn(1) * 4
    
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjMaintChainWheel As Object
    Dim ObjMaintSpokeCyl1 As Object
    Dim ObjMaintSpokeCyl2 As Object
    Dim ObjMaintGuideBaseCyl As Object
    Dim ObjMaintChain As Object
    Dim ObjMaintLeftGuide As Object
    Dim ObjMaintRightGuide  As Object
    Dim ObjMaintLeftGuideCyl As Object
    Dim ObjMaintRightGuideCyl As Object
    Dim ObjMaintLeftGuideCyl1 As Object
    Dim ObjMaintRightGuideCyl1 As Object
    
    Dim parChainWheelDia As Double
    Dim parOffset As Double
    Dim parFloorToValveDistance As Double
    Dim parOperatorHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parChainWheelDia = arrayOfInputs(2)
    parOffset = arrayOfInputs(3)
    parFloorToValveDistance = arrayOfInputs(4)
    parOperatorHeight = arrayOfInputs(5)
   
    iOutput = 0

 ' Insert your code for output 12(Maintenance ChainWheel )
   ''Chain Wheel is made by revolving a line string(Cross section of Chain wheel).
   
   ''ASSUMPTIONS
   'thickness of Chain Wheel is in Y Axis and is assumed as 25mm.(and in X  Direction too)
   'Height of the Guide Base cylinder as 50mm.
   'Guide base cylinder extends upto center of the Chainwheel.
   
   Dim dWheelThickness As Double
   Dim dGuideBaseCylHt As Double
   
   dWheelThickness = 0.025
   dGuideBaseCylHt = 0.07
   
   Dim oGeomFactory As IngrGeom3D.GeometryFactory
   Dim oLineStr As IngrGeom3D.LineString3d
   Dim oAxisVec As New AutoMath.DVector
   Dim oCenterPt As New AutoMath.DPosition
   Dim stPoint As New AutoMath.DPosition
   Dim enPoint As New AutoMath.DPosition
   Dim dAngle As Double
   
   Set oGeomFactory = New IngrGeom3D.GeometryFactory
   
   Dim oPoints(0 To 14) As Double
   
   oPoints(0) = 0
   oPoints(1) = parOperatorHeight - dGuideBaseCylHt - dWheelThickness / 2
   oPoints(2) = parChainWheelDia / 2
   
   oPoints(3) = 0
   oPoints(4) = parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2
   oPoints(5) = parChainWheelDia / 2
   
   oPoints(6) = 0
   oPoints(7) = parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2
   oPoints(8) = parChainWheelDia / 2 - dWheelThickness
   
   oPoints(9) = 0
   oPoints(10) = parOperatorHeight - dGuideBaseCylHt - dWheelThickness / 2
   oPoints(11) = parChainWheelDia / 2 - dWheelThickness
   
   oPoints(12) = oPoints(0)
   oPoints(13) = oPoints(1)
   oPoints(14) = oPoints(2)

   Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, oPoints)
   
   dAngle = 2 * PI
   oAxisVec.Set 0, 1, 0
   oCenterPt.Set 0, parOperatorHeight - dGuideBaseCylHt, 0
   
   Set ObjMaintChainWheel = PlaceRevolution(m_OutputColl, oLineStr, oAxisVec, oCenterPt, dAngle, True)
   
'' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintChainWheel
    Set ObjMaintChainWheel = Nothing

 ' Insert your code for output 13(Maintenance Spoke Cylinder 1)
   '' Assumed that spoke diameter is 20mm

   Dim dSpokeDia As Double
   dSpokeDia = 0.02

   stPoint.Set 0, parOperatorHeight - dGuideBaseCylHt, parChainWheelDia / 2 - dWheelThickness / 2
   enPoint.Set 0, parOperatorHeight - dGuideBaseCylHt, -parChainWheelDia / 2 + dWheelThickness / 2

   Set ObjMaintSpokeCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dSpokeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintSpokeCyl1
    Set ObjMaintSpokeCyl1 = Nothing

 ' Insert your code for output 14(Maintenance Spoke Cylinder 2)
   stPoint.Set parChainWheelDia / 2 - dWheelThickness / 2, parOperatorHeight - dGuideBaseCylHt, 0
   enPoint.Set -parChainWheelDia / 2 + dWheelThickness / 2, parOperatorHeight - dGuideBaseCylHt, 0

   Set ObjMaintSpokeCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dSpokeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintSpokeCyl2
    Set ObjMaintSpokeCyl2 = Nothing

 ' Insert your code for output 15(Maintenance Guide Base Cylinder)
   stPoint.Set 0, parOperatorHeight - dGuideBaseCylHt - dWheelThickness / 2, 0
   enPoint.Set 0, parOperatorHeight, 0

   Set ObjMaintGuideBaseCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parChainWheelDia / 3, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintGuideBaseCyl
    Set ObjMaintGuideBaseCyl = Nothing
'
'' Insert your code for output 16(Maintenance Chain)
   Dim oLine As IngrGeom3D.Line3d
   Dim oArc As IngrGeom3D.Arc3d
   Dim oComplexStr As IngrGeom3D.ComplexString3d
   Dim oCollection As New Collection
   
   Dim dChainPts(0 To 11) As Double
   Dim dChainWidth As Double
   Dim dChainThickness As Double
   Dim dSplit As Double
   
   ''Assumed ChainWidth and Thickness are 25mm
   dChainWidth = 0.025
   dChainThickness = 0.025
   dSplit = 0.00001 ''this is a split to form a chian with a complex string Projection.
   
   ''Chain Points.
   dChainPts(0) = -(parChainWheelDia / 2 + dChainWidth)
   dChainPts(1) = parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 '''''''''''''''''
   dChainPts(2) = 0

   dChainPts(3) = -(parChainWheelDia / 2 + dChainWidth)
   dChainPts(4) = parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2
   dChainPts(5) = -(parFloorToValveDistance - parOffset - parChainWheelDia / 2 - dChainWidth)

   dChainPts(6) = 0
   dChainPts(7) = parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2
   dChainPts(8) = -(parFloorToValveDistance - parOffset)
   ''Point on chain on the top of the chain wheel
   dChainPts(9) = 0
   dChainPts(10) = parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2
   dChainPts(11) = parChainWheelDia / 2 + dChainWidth

   ''Chain Outer string
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            dChainPts(0), dChainPts(1), dChainPts(2) - dSplit, _
                                            dChainPts(3), dChainPts(4), dChainPts(5))

   oCollection.Add oLine

   Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                            dChainPts(3), dChainPts(4), dChainPts(5), _
                                            dChainPts(6), dChainPts(7), dChainPts(8), _
                                            -dChainPts(3), dChainPts(4), dChainPts(5))

   oCollection.Add oArc

   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            -dChainPts(3), dChainPts(4), dChainPts(5), _
                                            -dChainPts(0), dChainPts(1), dChainPts(2))

   oCollection.Add oLine

   Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                            -dChainPts(0), dChainPts(1), dChainPts(2), _
                                            dChainPts(9), dChainPts(10), dChainPts(11), _
                                            dChainPts(0), dChainPts(1), dChainPts(2))

   oCollection.Add oArc

   ''Chain Innerstring
   '' inner Complex string points of chain are nothing but subtracting dChainWidth
   '' from outer string points.
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            dChainPts(0), dChainPts(1), dChainPts(2), _
                                            dChainPts(0) + dChainWidth, dChainPts(1), dChainPts(2))

   oCollection.Add oLine
   
   
   Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                            dChainPts(0) + dChainWidth, dChainPts(1), dChainPts(2), _
                                            dChainPts(9), dChainPts(10), dChainPts(11) - dChainWidth, _
                                            -dChainPts(0) - dChainWidth, dChainPts(1), dChainPts(2))

   oCollection.Add oArc
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        -dChainPts(0) - dChainWidth, dChainPts(1), dChainPts(2), _
                                        -dChainPts(3) - dChainWidth, dChainPts(4), dChainPts(5))

   oCollection.Add oLine
   
   Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                            -dChainPts(3) - dChainWidth, dChainPts(4), dChainPts(5), _
                                            dChainPts(6), dChainPts(7), dChainPts(8) + dChainWidth, _
                                            dChainPts(3) + dChainWidth, dChainPts(4), dChainPts(5))

   oCollection.Add oArc

   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dChainPts(3) + dChainWidth, dChainPts(4), dChainPts(5), _
                                dChainPts(0) + dChainWidth, dChainPts(1), dChainPts(2) - dSplit)

   oCollection.Add oLine
   
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dChainPts(0) + dChainWidth, dChainPts(1), dChainPts(2) - dSplit, _
                                dChainPts(0), dChainPts(1), dChainPts(2) - dSplit)

   oCollection.Add oLine
   

   stPoint.Set dChainPts(0), dChainPts(1), dChainPts(2)
   Set oComplexStr = PlaceTrCString(stPoint, oCollection)
   
   oAxisVec.Set 0, -1, 0
   Set ObjMaintChain = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, dChainThickness, True)
'
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintChain
    Set ObjMaintChain = Nothing
    
   Dim iCount As Integer
   For iCount = 1 To oCollection.Count
    oCollection.Remove 1
   Next iCount


'  Insert your code for output 17(Maintenance Left side Guide)
  ''Assumed Guide makes an angle 45deg with nozzle axis.

    Dim dGuideDia As Double
    dGuideDia = 0.01 ''Assumed Guide Dia as 10mm

    stPoint.Set 0, parOperatorHeight - dGuideBaseCylHt / 3, 0
    enPoint.Set -(parChainWheelDia / 2 + dChainWidth / 2) - dGuideDia, _
                                   parOperatorHeight - dGuideBaseCylHt / 3, _
                                   -(parChainWheelDia / 2 + dChainWidth / 2) - dGuideDia
                                   
   Set ObjMaintLeftGuide = PlaceCylinder(m_OutputColl, stPoint, enPoint, dGuideDia, True)
'' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintLeftGuide
    Set ObjMaintLeftGuide = Nothing

' ' Insert your code for output 18(Maintenance Right side Guide)
    stPoint.Set 0, parOperatorHeight - dGuideBaseCylHt / 3, 0
    enPoint.Set parChainWheelDia / 2 + dChainWidth / 2 + dGuideDia, _
                                    parOperatorHeight - dGuideBaseCylHt / 3, _
                                    -(parChainWheelDia / 2 + dChainWidth / 2) - dGuideDia

    Set ObjMaintRightGuide = PlaceCylinder(m_OutputColl, stPoint, enPoint, dGuideDia, True)
'' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintRightGuide
    Set ObjMaintRightGuide = Nothing

' ' Insert your code for output 19(Maintenance Lefft Guide Cylinder)
    ''Assumed guide cylinder is coming 2mm out of  the chain on each sides.
    stPoint.Set -(parChainWheelDia / 2 - 0.002), _
                         parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 + dGuideDia, _
                         -(parChainWheelDia / 2 + dChainWidth / 2)
                                
    enPoint.Set -(parChainWheelDia / 2 + dChainWidth + 0.002), _
                         parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 + dGuideDia, _
                         -(parChainWheelDia / 2 + dChainWidth / 2)

   Set ObjMaintLeftGuideCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * dGuideDia, True)
'' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintLeftGuideCyl
    Set ObjMaintLeftGuideCyl = Nothing

 ' Insert your code for output 20(Maintenance Right Guide Cylinder)
   ''Assumed guide cylinder is coming 2mm out of  the chain on each sides.
    stPoint.Set (parChainWheelDia / 2 - 0.002), _
                          parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 + dGuideDia, _
                          -(parChainWheelDia / 2 + dChainWidth / 2)
    enPoint.Set (parChainWheelDia / 2 + dChainWidth + 0.002), _
                          parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 + dGuideDia, _
                          -(parChainWheelDia / 2 + dChainWidth / 2)

   Set ObjMaintRightGuideCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * dGuideDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintRightGuideCyl
    Set ObjMaintRightGuideCyl = Nothing


' ' Insert your code for output 21(Maintenance Left Connecting Cylinder of Guide and Guide Cyl)
    stPoint.Set -(parChainWheelDia / 2 + dChainWidth / 2), _
                                    parOperatorHeight - dGuideBaseCylHt / 3, _
                                    -(parChainWheelDia / 2 + dChainWidth / 2)

    enPoint.Set -(parChainWheelDia / 2 + dChainWidth / 2), _
                      parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 + dGuideDia, _
                      -(parChainWheelDia / 2 + dChainWidth / 2)

   Set ObjMaintLeftGuideCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dGuideDia / 2, True)
'' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintLeftGuideCyl1
    Set ObjMaintLeftGuideCyl1 = Nothing

 ' Insert your code for output 22(Maintenance Right Connecting Cylinder of Guide and Guide Cyl)
    stPoint.Set parChainWheelDia / 2 + dChainWidth / 2, _
                                   parOperatorHeight - dGuideBaseCylHt / 3, _
                                   -(parChainWheelDia / 2 + dChainWidth / 2)
                                   
    enPoint.Set parChainWheelDia / 2 + dChainWidth / 2, _
                        parOperatorHeight - dGuideBaseCylHt + dWheelThickness / 2 + dGuideDia, _
                        -(parChainWheelDia / 2 + dChainWidth / 2)

   Set ObjMaintRightGuideCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dGuideDia / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintRightGuideCyl1
    Set ObjMaintRightGuideCyl1 = Nothing
    
    ''Removing Complex string
    Dim oComplexStrRem As IJDObject
    Set oComplexStrRem = oComplexStr
    oComplexStrRem.Remove
    Set oComplexStrRem = Nothing
    
    ''Removing Line string
    Dim oLineStrRem As IJDObject
    Set oLineStrRem = oComplexStr
    oLineStrRem.Remove
    Set oLineStrRem = Nothing

    Set oGeomFactory = Nothing
    Set oArc = Nothing
    Set oLine = Nothing
    Set oComplexStr = Nothing
    Set oLineStr = Nothing
    Set oCollection = Nothing
    Set oAxisVec = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oCenterPt = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
