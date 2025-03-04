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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Mar 19 2007
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       90 Degree Expanding/Reducing Cable Tray Bend (Sourec: 'electrical Model (1).pdf')
'
'   Change History:
'   dd.mmm.yyyy     who                    change description
'   -----------         -----                ------------------
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjHoriTangent  As Object
    Dim ObjInclTangent  As Object
    
    Dim parCTBendThroatRadius As Double
    Dim parCTBendOuterRadius As Double
    Dim parTangentLength As Double
    Dim oPort1 As AutoMath.DPosition  'Port 1 center point
    Dim oPort2 As AutoMath.DPosition  'Port 2 center point
    Dim oProjVector As AutoMath.DVector
    Dim oLineStrCP As AutoMath.DPosition
    Set oPort1 = New AutoMath.DPosition
    Set oPort2 = New AutoMath.DPosition
    Set oProjVector = New AutoMath.DVector
    Set oLineStrCP = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCTBendThroatRadius = arrayOfInputs(2)
    parCTBendOuterRadius = arrayOfInputs(3)
    
    iOutput = 0

    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    parTangentLength = oTrayPart.TangentLength
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If parTangentLength = 0 Then parTangentLength = 0.000001
    
'   Get port 1 properties
    Dim dActualWidth1 As Double
    Dim dActualDepth1 As Double
    Call RetrieveCableTrayPortProperties(1, oPartFclt, dActualWidth1, dActualDepth1)
    
'   Get port 2 properties
    Dim dActualWidth2 As Double
    Dim dActualDepth2 As Double
    Call RetrieveCableTrayPortProperties(2, oPartFclt, dActualWidth2, dActualDepth2)

    Dim dFace2toCenter As Double
    dFace2toCenter = dActualWidth1 / 2 + parCTBendThroatRadius + parTangentLength

' Insert your code for output 1(U Shape Horizontal Tangent)
    Dim CP As AutoMath.DPosition 'arc center point
    Set CP = New AutoMath.DPosition
    Dim HD              As Double
    Dim HW              As Double
    Dim LineStrPoints(0 To 11)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    HD = dActualDepth1 / 2
    HW = dActualWidth1 / 2
    Dim dFace1toCenter As Double
    dFace1toCenter = dActualWidth2 / 2 + parCTBendThroatRadius + parTangentLength

    oPort1.Set -dFace1toCenter, 0, 0
    
    LineStrPoints(0) = oPort1.x
    LineStrPoints(1) = oPort1.y + HW
    LineStrPoints(2) = oPort1.z + HD
    
    LineStrPoints(3) = oPort1.x
    LineStrPoints(4) = oPort1.y + HW
    LineStrPoints(5) = oPort1.z - HD
    
    LineStrPoints(6) = oPort1.x
    LineStrPoints(7) = oPort1.y - HW
    LineStrPoints(8) = oPort1.z - HD
    
    LineStrPoints(9) = oPort1.x
    LineStrPoints(10) = oPort1.y - HW
    LineStrPoints(11) = oPort1.z + HD
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    oProjVector.Set 1, 0, 0
    Set ObjHoriTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, parTangentLength, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriTangent
    Set ObjHoriTangent = Nothing
    Set oLineString = Nothing
    
 ' Insert your code for output 2(Bend Inner Side)
    Dim objBIA  As IngrGeom3D.Arc3d
    Dim oSBA   As AutoMath.DPosition
    Dim oEBA   As AutoMath.DPosition
    Dim oCBA   As AutoMath.DPosition
    
    Set oSBA = New AutoMath.DPosition
    Set oEBA = New AutoMath.DPosition
    Set oCBA = New AutoMath.DPosition

    oSBA.Set -dActualWidth2 / 2, dActualWidth1 / 2 + parCTBendThroatRadius, -dActualDepth1 / 2
    oEBA.Set -(dActualWidth2 / 2 + parCTBendThroatRadius), dActualWidth1 / 2, -dActualDepth1 / 2
    oCBA.Set oEBA.x, oSBA.y, -dActualDepth1 / 2
    Set objBIA = PlaceTrArcByCenter(oSBA, oEBA, oCBA)
    
    oProjVector.Set 0, 0, 1

    Dim ObjBendInnerSide As IngrGeom3D.Projection3d
    Set ObjBendInnerSide = PlaceProjection(m_OutputColl, objBIA, oProjVector, dActualDepth1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBendInnerSide
    Set ObjBendInnerSide = Nothing

' Insert your code for output 3(Bend Outer Side)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    'Construct branch left tangent line at outer edge
    stPoint.Set -(dActualWidth2 / 2 + parCTBendThroatRadius), -dActualWidth1 / 2, -dActualDepth1 / 2
    enPoint.Set (dActualWidth2 / 2 - parCTBendOuterRadius), stPoint.y, stPoint.z
    Dim oBLTLine As IngrGeom3D.Line3d
    Set oBLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
          stPoint.x, stPoint.y, stPoint.z, enPoint.x, enPoint.y, enPoint.z)

    'Construct branch arc at outer edge
    oSBA.Set (dActualWidth2 / 2 - parCTBendOuterRadius), -dActualWidth1 / 2, -dActualDepth1 / 2
    oEBA.Set dActualWidth2 / 2, -dActualWidth1 / 2 + parCTBendOuterRadius, -dActualDepth1 / 2
    oCBA.Set oSBA.x, oEBA.y, -dActualDepth1 / 2
    Dim objBOA  As IngrGeom3D.Arc3d
    Set objBOA = PlaceTrArcByCenter(oSBA, oEBA, oCBA)
    
    'Construct branch upward tangent line at outer edge
    Dim oBUTLine As IngrGeom3D.Line3d
    stPoint.Set dActualWidth2 / 2, -dActualWidth1 / 2 + parCTBendOuterRadius, -dActualDepth1 / 2
    enPoint.Set stPoint.x, dFace2toCenter - parTangentLength, stPoint.z
    Set oBUTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
          stPoint.x, stPoint.y, stPoint.z, enPoint.x, enPoint.y, enPoint.z)
          
    Dim obranchcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Set obranchcurves = New Collection
    obranchcurves.Add oBLTLine
    obranchcurves.Add objBOA
    obranchcurves.Add oBUTLine
    
    Dim StartBC   As AutoMath.DPosition
    Set StartBC = New AutoMath.DPosition
    StartBC.Set -(dActualWidth2 / 2 + parCTBendThroatRadius), -dActualWidth1 / 2, -dActualDepth1 / 2
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    
    Dim ObjBendOuterSide As IngrGeom3D.Projection3d
    Set ObjBendOuterSide = PlaceProjection(m_OutputColl, objHBcurves, oProjVector, dActualDepth1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBendOuterSide
    Set ObjBendOuterSide = Nothing

' Insert your code for output 4(Bend Bottom)
    'Construct line joining outer up tangent end point to inner arc start point
    stPoint.Set dActualWidth2 / 2, dActualWidth1 / 2 + parCTBendThroatRadius, -dActualDepth1 / 2
    enPoint.Set -dActualWidth2 / 2, stPoint.y, stPoint.z
    
    Dim oOILine As IngrGeom3D.Line3d
    Set oOILine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
          stPoint.x, stPoint.y, stPoint.z, enPoint.x, enPoint.y, enPoint.z)

    obranchcurves.Add oOILine
    obranchcurves.Add objBIA

    'Construct line joining inner arc end point to outer bottom tangent start point
    stPoint.Set -dActualWidth2 / 2 - parCTBendThroatRadius, dActualWidth1 / 2, -dActualDepth1 / 2
    enPoint.Set stPoint.x, -dActualWidth1 / 2, stPoint.z
    
    Dim oIOLine As IngrGeom3D.Line3d
    Set oIOLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
          stPoint.x, stPoint.y, stPoint.z, enPoint.x, enPoint.y, enPoint.z)
    obranchcurves.Add oIOLine
    
    StartBC.Set -(dActualWidth2 / 2 + parCTBendThroatRadius), -dActualWidth1 / 2, -dActualDepth1 / 2
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    
    Dim TrayBottom As IngrGeom3D.Plane3d
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    Set TrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
        -(dActualWidth2 / 2 + parCTBendThroatRadius), -dActualWidth1 / 2, -dActualDepth1 / 2, oDirProj.x, oDirProj.y, oDirProj.z)
        
    oDirProj.Set 0, 0, -1
    Call TrayBottom.AddBoundary(objHBcurves)
                    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), TrayBottom
    Set TrayBottom = Nothing

' Insert your code for output 3(U Shape Inclined Tangent)
    Dim oInclinedLineStr As IngrGeom3D.LineString3d
    oPort2.Set 0, dFace2toCenter, 0
    
    HW = dActualWidth2 / 2
    HD = dActualDepth2 / 2
    LineStrPoints(0) = oPort2.x - HW
    LineStrPoints(1) = oPort2.y
    LineStrPoints(2) = oPort2.z + HD
    
    LineStrPoints(3) = oPort2.x - HW
    LineStrPoints(4) = oPort2.y
    LineStrPoints(5) = oPort2.z - HD
    
    LineStrPoints(6) = oPort2.x + HW
    LineStrPoints(7) = oPort2.y
    LineStrPoints(8) = oPort2.z - HD
    
    LineStrPoints(9) = oPort2.x + HW
    LineStrPoints(10) = oPort2.y
    LineStrPoints(11) = oPort2.z + HD
    Set oInclinedLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    oProjVector.Set 0, -1, 0
    Set ObjInclTangent = PlaceProjection(m_OutputColl, oInclinedLineStr, oProjVector, parTangentLength, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInclTangent
    Set ObjInclTangent = Nothing
    Set oProjVector = Nothing
    Set oInclinedLineStr = Nothing
    
' Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc

    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort1 = Nothing
    Set oPortLocation = Nothing

' Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition

    oDir.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing

    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set CP = Nothing
    Set oGeomFactory = Nothing
    Set oSBA = Nothing
    Set oEBA = Nothing
    Set oCBA = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set obranchcurves = Nothing
    Set StartBC = Nothing
    Set oDirProj = Nothing
    Set oTrayPart = Nothing
    Set objBIA = Nothing
    Set oBLTLine = Nothing
    Set objBOA = Nothing
    Set oBUTLine = Nothing
    Set objHBcurves = Nothing
    Set oOILine = Nothing
    Set oIOLine = Nothing
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
End Sub
