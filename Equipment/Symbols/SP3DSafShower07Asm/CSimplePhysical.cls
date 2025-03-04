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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-4:59:47 AM"
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
    
    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjSupport As Object
    Dim ObjSpine As Object
    Dim ObjElbow As Object
    Dim ObjArm As Object
    Dim ObjWrist As Object
    Dim ObjNeck As Object
    Dim ObjHead As Object
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    Dim ObjBox4 As Object
    Dim ObjBox5 As Object
    Dim ObjBox6 As Object
    Dim ObjBox7 As Object
   

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0

' Insert your code for output 1(Support Cone)
    strOrigin = "W 0.356 S 0.356 U 0.197"
    strDirection = "N 0 U 90"
    dblDiameter = 0.229
    dblHeight = 0.076
    Set ObjSupport = CreateCone(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport
    Set ObjSupport = Nothing
    
' Insert your code for output 2(Spine Cylinder)
    strOrigin = "W 0.356 S 0.356 U 1.3"
    dblDiameter = 0.055
    dblLength = 2.283
    strDirection = "N 0 U 90"
    Set ObjSpine = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSpine
    Set ObjSpine = Nothing

' Insert your code for output 3(Elbow CTor)
    strOrigin = "W 0.326 S 0.326 U 2.442"
    dblInsideRadius = 0.015
    dblOutsideRadius = 0.07
    strXdir = "NE 45 U 0"
    strYdir = "N 0 D -90"
    dblSweepAngle = 90
    Set ObjElbow = CreateCircularTorus(m_outputColl, strOrigin, dblInsideRadius, dblOutsideRadius, , strXdir, _
                    strYdir, dblSweepAngle)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing

' Insert your code for output 4(Arm Cylinder)
    strOrigin = "W 0.178 S 0.178 U 2.485"
    dblDiameter = 0.055
    dblLength = 0.419
    strDirection = "NE 45 U 0"
    Set ObjArm = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjArm
    Set ObjArm = Nothing

' Insert your code for output 5(Wrist CTor)
    strOrigin = "W 0.030 S 0.030 U 2.443"
    dblInsideRadius = 0.015
    dblOutsideRadius = 0.07
    strXdir = "SW 225 U 0"
    strYdir = "N 0 D -90"
    dblSweepAngle = 90
    Set ObjWrist = CreateCircularTorus(m_outputColl, strOrigin, dblInsideRadius, dblOutsideRadius, , strXdir, _
                    strYdir, dblSweepAngle)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjWrist
    Set ObjWrist = Nothing

' Insert your code for output 6(Neck Cylinder)
    strOrigin = "E 0 N 0 U 2.415"
    dblDiameter = 0.055
    dblLength = 0.055
    strDirection = "N 0 D -90"
    Set ObjNeck = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjNeck
    Set ObjNeck = Nothing

' Insert your code for output 7(Head Dish)
    strOrigin = "E 0 N 0 U 2.317"
    strDirection = "N 0 U 90"
    dblDiameter = 0.305
    dblHeight = 0.076
    dblKnuckleRadius = 0
    Set ObjHead = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, _
                    True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjHead
    Set ObjHead = Nothing

' Insert your code for output 8(Box 1)
    strOrigin = "E 0 N 0 U 0.013"
    dblXLength = 1.22
    dblYLength = 1.322
    dblZLength = 0.025
    Set oriOrientation = New Orientation
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutZ = 180
    oriOrientation.ApplyRotations
    Set ObjBox1 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    Set oriOrientation = Nothing

' Insert your code for output 9(Box 2)
    strOrigin = "E 0 S 0.648 U 1.296"
    dblXLength = 1.22
    dblYLength = 0.025
    dblZLength = 2.591
    Set ObjBox2 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing

' Insert your code for output 10(Box 3)
    strOrigin = "E 0.598 N 0 U 1.296"
    dblXLength = 0.025
    dblYLength = 1.321
    dblZLength = 2.591
    Set ObjBox3 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing

' Insert your code for output 11(Box 4)
    strOrigin = "W 0.598 N 0 U 1.296"
    dblXLength = 0.025
    dblYLength = 1.321
    dblZLength = 2.591
    Set ObjBox4 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox4
    Set ObjBox4 = Nothing

' Insert your code for output 12(Box 5)
    strOrigin = "E 0 N 0.648 U 1.296"
    dblXLength = 1.22
    dblYLength = 0.025
    dblZLength = 2.591
    Set ObjBox5 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox5
    Set ObjBox5 = Nothing

' Insert your code for output 13(Box 6)
    strOrigin = "W 0.0262 N 0.661 U 1.259"
    dblXLength = 0.53
    dblYLength = 0.03
    dblZLength = 2.2
    Set ObjBox6 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox6
    Set ObjBox6 = Nothing

' Insert your code for output 14(Box 7)
    strOrigin = "E 0.270 N 0.661 U 1.259"
    dblXLength = 0.53
    dblYLength = 0.03
    dblZLength = 2.2
    Set ObjBox7 = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox7
    Set ObjBox7 = Nothing


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
    startPos.Set -0.838, -0.356, 0.324
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set -1, 0, 0
    iDistribPort.SetDirectionVector Axis

    Axis.Set 0, 1, 0
    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.482
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
    startPos.Set -0.838, -0.356, 2.308
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set -1, 0, 0
    iDistribPort.SetDirectionVector Axis

    Axis.Set 0, 1, 0
    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.482
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
    startPos.Set 0.65, 0, 0.064
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set 1, 0, 0
    iDistribPort.SetDirectionVector Axis

    Axis.Set 0, 1, 0
    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.08
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


