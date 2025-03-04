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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 19, Apr 2006
'   Description:
'    This symbol is Pipe Sleeve with Puddle Flange (Details taken from
'    product database at www.doyma.de, Curaflex Pipe Sleeve 9000 Technical Information, Doyma)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
      PI = 4 * Atn(1)
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjLHSFlowSensor As Object
    Dim ObjLHSBaseforFlowSensor As Object
    Dim ObjLHSClampforFlowSensor As Object
    Dim ObjLHSClampBolt As Object
    Dim ObjRHSFlowSensor As Object
    Dim ObjRHSBaseforFlowSensor As Object
    Dim ObjRHSClampforFlowSensor As Object
    Dim ObjRHSClampBolt As Object

    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parSensorSpacing As Double
    Dim parClampSpacing As Double
    Dim parClamptoEnd As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentWidth = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    parInstrumentDiameter = arrayOfInputs(4)
    parSensorSpacing = arrayOfInputs(5)
    parClampSpacing = arrayOfInputs(6)
    parClamptoEnd = arrayOfInputs(7)

    iOutput = 0

'   Origin is assumed to be at the mid-point of two sensors and on the pipe axis

'   Insert your code for output 1(L H S Flow Sensor)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

'   Assumptions:
'   1. Sensor base is assumed to be a curved surface to have contact with the pipe surface. So
'      the sensor base is approximated with a cylinder shown upto pipeDiam/6
'   2. Sensor top diameter is taken to be 50% of the Sensor base diameter (parInstrumentDiameter)
'   3. Sensor top is considered to be cylinder and it's height is equal to 40% of parInstrumentHeight
    Dim dSensorStPointHeight As Double
    Dim dSensorTopDiameter As Double
    Dim dHeight As Double

    dSensorStPointHeight = pipeDiam / 6
    dSensorTopDiameter = parInstrumentDiameter / 2
    dHeight = 0.4 * parInstrumentHeight
    
    Dim dPoints(0 To 26) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    
    oCenterPoint.Set 0, 0, 0

'   Starting point is taken at Sensor top
    dPoints(0) = oCenterPoint.x - parSensorSpacing / 2 - parInstrumentDiameter / 2
    dPoints(1) = oCenterPoint.y + pipeDiam / 2 + parInstrumentHeight
    dPoints(2) = oCenterPoint.z  'remains same for other points
        
    dPoints(3) = dPoints(0) - dSensorTopDiameter / 2
    dPoints(4) = dPoints(1)
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = dPoints(3)
    dPoints(7) = dPoints(4) - dHeight
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = oCenterPoint.x - parSensorSpacing / 2 - parInstrumentDiameter
    dPoints(10) = oCenterPoint.y + pipeDiam / 2 + dHeight
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(9)
    dPoints(13) = oCenterPoint.y + pipeDiam / 2 - dSensorStPointHeight
    dPoints(14) = oCenterPoint.z
        
    dPoints(15) = dPoints(12) + parInstrumentDiameter / 2
    dPoints(16) = dPoints(13)
    dPoints(17) = oCenterPoint.z

    Dim oFlowSensorCenPt As AutoMath.DPosition
    Set oFlowSensorCenPt = New AutoMath.DPosition
    oFlowSensorCenPt.Set oCenterPoint.x - parSensorSpacing / 2 - parInstrumentDiameter / 2, _
                     oCenterPoint.y, oCenterPoint.z
                     
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim axisVect    As AutoMath.DVector
    Set axisVect = New AutoMath.DVector

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dPoints)
    axisVect.Set 0, 1, 0
    Set ObjLHSFlowSensor = PlaceRevolution(m_OutputColl, oLineString, axisVect, oFlowSensorCenPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSFlowSensor
    Set ObjLHSFlowSensor = Nothing

'   Insert your code for output 2(L H S Base for Flow Sensor)
'   Assumptions:
'    1. The base that supports the sensor is assumed to be a box with height of 10% of parInstrumentHeight
'       measured above the pipe surface
'    2. The base is taken to overlap in to pipe surface by dSensorStPointHeight
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -parClampSpacing / 2, pipeDiam / 2 - dSensorStPointHeight, -parInstrumentWidth / 2
    enPoint.Set -parSensorSpacing / 2 - parInstrumentDiameter / 2, _
                pipeDiam / 2 + 0.1 * parInstrumentHeight, parInstrumentWidth / 2

    Set ObjLHSBaseforFlowSensor = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSBaseforFlowSensor
    Set ObjLHSBaseforFlowSensor = Nothing

'   Insert your code for output 3(L H S Clamp for Flow Sensor)
'   Assumptions:
'    1. Clamp Thickness if taken to be 10 mm
'    2. Clamp strap width is taken to be 25% of parClamptoEnd value
    Dim dClampThk As Double
    Dim dClampStrapWidth As Double

    dClampThk = 0.01
    dClampStrapWidth = 0.25 * parClamptoEnd

    dPoints(0) = oCenterPoint.x - parClampSpacing / 2 + parClamptoEnd - dClampStrapWidth / 2
    dPoints(1) = oCenterPoint.y + pipeDiam / 2
    dPoints(2) = oCenterPoint.z
        
    dPoints(3) = dPoints(0) + dClampStrapWidth
    dPoints(4) = dPoints(1)
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = dPoints(3)
    dPoints(7) = dPoints(4) + dClampThk
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = dPoints(6) - dClampStrapWidth
    dPoints(10) = dPoints(7)
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(0)
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    axisVect.Set 1, 0, 0
    Set ObjLHSClampforFlowSensor = PlaceRevolution(m_OutputColl, oLineString, axisVect, oCenterPoint, 2 * PI, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSClampforFlowSensor
    Set ObjLHSClampforFlowSensor = Nothing

'   Insert your code for output 4(L H S Clamp Bolt)
'   Assumptions:
'    1. Height measured upto Bolt top is taken to be half of parInstrumentHeight
'    2. Total Bolt is shown as three pieces with heights 1:2:4 ratio with their diameters in 1:1.1:1.2 ratio
'       respectively.
    Dim dHtUptoBoltTop As Double
    Dim dSmallDia As Double
    Dim dMediumDia As Double
    Dim dLargeDia As Double
    dHtUptoBoltTop = 0.5 * parInstrumentHeight
    dSmallDia = 0.14 * parInstrumentDiameter
    dMediumDia = 1.1 * dSmallDia
    dLargeDia = 1.2 * dSmallDia

'   Starting point is taken at Bolt top
    dPoints(0) = oCenterPoint.x - parClampSpacing / 2 + parClamptoEnd
    dPoints(1) = oCenterPoint.y + pipeDiam / 2 + dHtUptoBoltTop
    dPoints(2) = oCenterPoint.z  'remains same for other points
        
    dPoints(3) = dPoints(0) - dSmallDia / 2
    dPoints(4) = dPoints(1)
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = dPoints(3)
    dPoints(7) = dPoints(4) - dHtUptoBoltTop / 7
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = oCenterPoint.x - parClampSpacing / 2 + parClamptoEnd - dMediumDia / 2
    dPoints(10) = dPoints(7)
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(9)
    dPoints(13) = dPoints(10) - 2 * dHtUptoBoltTop / 7
    dPoints(14) = oCenterPoint.z
        
    dPoints(15) = oCenterPoint.x - parClampSpacing / 2 + parClamptoEnd - dLargeDia / 2
    dPoints(16) = dPoints(13)
    dPoints(17) = oCenterPoint.z

    dPoints(18) = dPoints(15)
    dPoints(19) = pipeDiam / 2
    dPoints(20) = oCenterPoint.z
        
    dPoints(21) = oCenterPoint.x - parClampSpacing / 2 + parClamptoEnd
    dPoints(22) = dPoints(19)
    dPoints(23) = oCenterPoint.z
        
    Dim oBoltCenPt As AutoMath.DPosition
    Set oBoltCenPt = New AutoMath.DPosition
    oBoltCenPt.Set oCenterPoint.x - parClampSpacing / 2 + parClamptoEnd, _
                     oCenterPoint.y, oCenterPoint.z

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dPoints)
    axisVect.Set 0, 1, 0
    Set ObjLHSClampBolt = PlaceRevolution(m_OutputColl, oLineString, axisVect, oBoltCenPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSClampBolt
    Set ObjLHSClampBolt = Nothing
    
'   Insert your code for output 5(R H S Flow Sensor)
'   Starting point is taken at Sensor top
    dPoints(0) = oCenterPoint.x + parSensorSpacing / 2 + parInstrumentDiameter / 2
    dPoints(1) = oCenterPoint.y + pipeDiam / 2 + parInstrumentHeight
    dPoints(2) = oCenterPoint.z  'remains same for other points
        
    dPoints(3) = dPoints(0) - dSensorTopDiameter / 2
    dPoints(4) = dPoints(1)
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = dPoints(3)
    dPoints(7) = dPoints(4) - dHeight
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = oCenterPoint.x + parSensorSpacing / 2
    dPoints(10) = oCenterPoint.y + pipeDiam / 2 + dHeight
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(9)
    dPoints(13) = oCenterPoint.y + pipeDiam / 2 - dSensorStPointHeight
    dPoints(14) = oCenterPoint.z
        
    dPoints(15) = dPoints(12) + parInstrumentDiameter / 2
    dPoints(16) = dPoints(13)
    dPoints(17) = oCenterPoint.z

    oFlowSensorCenPt.Set oCenterPoint.x + parSensorSpacing / 2 + parInstrumentDiameter / 2, _
                     oCenterPoint.y, oCenterPoint.z
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dPoints)
    axisVect.Set 0, 1, 0
    Set ObjRHSFlowSensor = PlaceRevolution(m_OutputColl, oLineString, axisVect, oFlowSensorCenPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSFlowSensor
    Set ObjRHSFlowSensor = Nothing
    
'   Insert your code for output 6(R H S Base for Flow Sensor)
    stPoint.Set parSensorSpacing / 2 + parInstrumentDiameter / 2, _
                pipeDiam / 2 - dSensorStPointHeight, -parInstrumentWidth / 2
    enPoint.Set parClampSpacing / 2, pipeDiam / 2 + 0.1 * parInstrumentHeight, parInstrumentWidth / 2

    Set ObjRHSBaseforFlowSensor = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSBaseforFlowSensor
    Set ObjRHSBaseforFlowSensor = Nothing

'   Insert your code for output 7(R H S Clamp for Flow Sensor)
    dPoints(0) = oCenterPoint.x + parClampSpacing / 2 - parClamptoEnd - dClampStrapWidth / 2
    dPoints(1) = oCenterPoint.y + pipeDiam / 2
    dPoints(2) = oCenterPoint.z
        
    dPoints(3) = dPoints(0) + dClampStrapWidth
    dPoints(4) = dPoints(1)
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = dPoints(3)
    dPoints(7) = dPoints(4) + dClampThk
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = dPoints(6) - dClampStrapWidth
    dPoints(10) = dPoints(7)
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(0)
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    axisVect.Set 1, 0, 0
    Set ObjRHSClampforFlowSensor = PlaceRevolution(m_OutputColl, oLineString, axisVect, oCenterPoint, 2 * PI, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSClampforFlowSensor
    Set ObjRHSClampforFlowSensor = Nothing
    
'   Insert your code for output 8(R H S Clamp Bolt)
'   Starting point is taken at Bolt top
    dPoints(0) = oCenterPoint.x + parClampSpacing / 2 - parClamptoEnd
    dPoints(1) = oCenterPoint.y + pipeDiam / 2 + dHtUptoBoltTop
    dPoints(2) = oCenterPoint.z  'remains same for other points
        
    dPoints(3) = dPoints(0) - dSmallDia / 2
    dPoints(4) = dPoints(1)
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = dPoints(3)
    dPoints(7) = dPoints(4) - dHtUptoBoltTop / 7
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = oCenterPoint.x + parClampSpacing / 2 - parClamptoEnd - dMediumDia / 2
    dPoints(10) = dPoints(7)
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(9)
    dPoints(13) = dPoints(10) - 2 * dHtUptoBoltTop / 7
    dPoints(14) = oCenterPoint.z
        
    dPoints(15) = oCenterPoint.x + parClampSpacing / 2 - parClamptoEnd - dLargeDia / 2
    dPoints(16) = dPoints(13)
    dPoints(17) = oCenterPoint.z

    dPoints(18) = dPoints(15)
    dPoints(19) = pipeDiam / 2
    dPoints(20) = oCenterPoint.z
        
    dPoints(21) = oCenterPoint.x + parClampSpacing / 2 - parClamptoEnd
    dPoints(22) = dPoints(19)
    dPoints(23) = oCenterPoint.z

    oBoltCenPt.Set oCenterPoint.x + parClampSpacing / 2 - parClamptoEnd, _
                     oCenterPoint.y, oCenterPoint.z
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dPoints)
    axisVect.Set 0, 1, 0
    Set ObjRHSClampBolt = PlaceRevolution(m_OutputColl, oLineString, axisVect, oBoltCenPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSClampBolt
    Set ObjRHSClampBolt = Nothing
    Set geomFactory = Nothing
    Set axisVect = Nothing
    Set oLineString = Nothing
    Set oFlowSensorCenPt = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oBoltCenPt = Nothing
    Set oCenterPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

