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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:06:24 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Feb 24 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim strOrigin As String
    Dim strDirection As String
    Dim dblDiameter As Double
    Dim dblHeight As Double
    Dim dblLength As Double
    Dim dblInsideRadius As Double
    Dim dblOutsideRadius As Double
    Dim strXdir As String
    Dim strYdir As String
    Dim dblSweepAngle As Double
    Dim dblKnuckleRadius As Double
    Dim oriOrientation As Orientation
    Dim dblXLength As Double
    Dim dblYLength As Double
    Dim dblZLength As Double
    
    Dim iOutput     As Double
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    Dim ObjBox4 As Object
    Dim ObjBox5 As Object
    Dim ObjCylinder1 As Object
    Dim ObjDish1 As Object
    Dim ObjCTor1 As Object
    Dim ObjCylinder2 As Object
    Dim ObjCTor2 As Object
    Dim ObjCylinder3 As Object
    Dim ObjDish2 As Object
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0

' Insert your code for output 1(Box1)
    strOrigin = "E 0 N 0 U 0.005"
    dblXLength = 1.01
    dblYLength = 1.1
    dblZLength = 0.01
    Set oriOrientation = New Orientation
    oriOrientation.ResetDefaultAxis
'    oriOrientation.RotationAboutZ = 180
'    oriOrientation.ApplyRotations
    Set ObjBox1 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
' Insert your code for output 2(Box2)
    strOrigin = "E 0.5 N 0 U 1.3"
    dblXLength = 0.01
    dblYLength = 1.1
    dblZLength = 2.6
    Set ObjBox2 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing
    
' Insert your code for output 3(Box3)
    strOrigin = "W 0.5 N 0 U 1.3"
    dblXLength = 0.01
    dblYLength = 1.1
    dblZLength = 2.6
    Set ObjBox3 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing
    
' Insert your code for output 4(Box4)
    strOrigin = "E 0 S 0.545 U 1.3"
    dblXLength = 1.01
    dblYLength = 0.01
    dblZLength = 2.6
    Set ObjBox4 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox4
    Set ObjBox4 = Nothing
    
' Insert your code for output 5(Box5)
    strOrigin = "E 0 N 0.545 U 2.450"
    dblXLength = 1.01
    dblYLength = 0.01
    dblZLength = 0.3
    Set ObjBox5 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox5
    Set ObjBox5 = Nothing
    
' Insert your code for output 6(Cylinder 1)
    strOrigin = "E 0 S 0.220 U 0.5"
    dblDiameter = 0.055
    dblLength = 0.8
    strDirection = "N 0 U 90"
    Set ObjCylinder1 = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
    
' Insert your code for output 7(Dish 1)
    strOrigin = "E 0 S 0.220 U 0.976"
    strDirection = "N 0 D -90"
    dblDiameter = 0.305
    dblHeight = 0.076
    dblKnuckleRadius = 0
    Set ObjDish1 = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, _
                    True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish1
    Set ObjDish1 = Nothing
    
' Insert your code for output 8(CTor1)
    strOrigin = "E 0 S 0.372 U 2.281"
    dblInsideRadius = 0.015
    dblOutsideRadius = 0.07
    strXdir = "N 0 U 0"
    strYdir = "N 0 D -90"
    dblSweepAngle = 90
    Set ObjCTor1 = CreateCircularTorus(m_outputColl, strOrigin, dblInsideRadius, dblOutsideRadius, , strXdir, _
                    strYdir, dblSweepAngle)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCTor1
    Set ObjCTor1 = Nothing
   
' Insert your code for output 9(Cylinder2)
    strOrigin = "E 0 S 0.097 U 2.323"
    dblDiameter = 0.055
    dblLength = 0.55
    strDirection = "S 180 U 0"
    Set ObjCylinder2 = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
    Set ObjCylinder2 = Nothing
    
' Insert your code for output 10(CTor2)
    strOrigin = "E 0 N 0.178 U 2.280"
    dblInsideRadius = 0.015
    dblOutsideRadius = 0.07
    strXdir = "S 180 U 0"
    strYdir = "N 0 D -90"
    dblSweepAngle = 90
    Set ObjCTor2 = CreateCircularTorus(m_outputColl, strOrigin, dblInsideRadius, dblOutsideRadius, , strXdir, _
                    strYdir, dblSweepAngle)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCTor2
    Set ObjCTor2 = Nothing
    
' Insert your code for output 11(Cylinder3)
    strOrigin = "E 0 N 0.220 U 2.240"
    dblDiameter = 0.055
    dblLength = 0.08
    strDirection = "N 0 U 0"
    Set ObjCylinder3 = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder3
    Set ObjCylinder3 = Nothing
    
' Insert your code for output 12(Dish2)
    strOrigin = "E 0 N 0.220 U 2.124"
    strDirection = "N 0 U 90"
    dblDiameter = 0.305
    dblHeight = 0.076
    dblKnuckleRadius = 0
    Set ObjDish2 = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, _
                    True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish2
    Set ObjDish2 = Nothing
    
    
'MsgBox "Begin Nozzle"
'' Place Nozzle 1

    Dim NozzlePHFactory As New NozzlePHFactory
    Dim oPipeNozzle As PipeNozzle
    Dim iNozzle As IJDNozzle
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim startPos As IJDPosition
    Dim Axis As IJDVector

    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 1, False, m_outputColl.ResourceManager)
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    Set startPos = New DPosition
    startPos.Set 0.205, -0.7, 2.2
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set 0, -1, 0
    iDistribPort.SetDirectionVector Axis

'    Axis.Set 1, 0, 0
'    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.17
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPipeNozzle
    Set oPipeNozzle = Nothing
    Set iNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set startPos = Nothing
    Set Axis = Nothing
    
'' Place Nozzle 2
 
    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 2, False, m_outputColl.ResourceManager)
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    Set startPos = New DPosition
    startPos.Set -0.365, -0.7, 2.2
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set 0, -1, 0
    iDistribPort.SetDirectionVector Axis

'    Axis.Set 1, 0, 0
'    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.17
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPipeNozzle
    Set oPipeNozzle = Nothing
    Set iNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set startPos = Nothing
    Set Axis = Nothing
        
'' Place Nozzle 3
 
    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 3, False, m_outputColl.ResourceManager)
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    Set startPos = New DPosition
    startPos.Set 0, -0.22, 0
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    
    Axis.Set 0, 0, -1
    iDistribPort.SetDirectionVector Axis

'    Axis.Set 0, 1, 0
'    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.17
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPipeNozzle
    Set oPipeNozzle = Nothing
    Set iNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set startPos = Nothing
    Set Axis = Nothing
    Set NozzlePHFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim cptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - cptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzlePH(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + cptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzlePH(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''



