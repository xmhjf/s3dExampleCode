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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:07:18 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Feb 21 2003
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
    
    Dim iOutput     As Double
    Dim ObjShowerBase As Object
    Dim ObjShowerVSupport As Object
    Dim ObjShowerHSupport As Object
    Dim ObjDishAttachSupport As Object
    
    Dim objCone As Object
    Dim ObjCTorus As Object
    Dim ObjDish As Object
    
    Dim strPosition As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    Dim dblHeight As Double
    Dim dblInRadius As Double
    Dim dblOutRadius As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0

 ' Insert your code for output 1(Shower Base)
    strPosition = "E 0 S 0.61 U 0.005"
    dblDiameter = 1.828
    dblLength = 0.01
 
    Set ObjShowerBase = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, "N 0 U 90")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerBase
    Set ObjShowerBase = Nothing
    
 ' Insert your code for output 2(Shower Vert. Support)
    strPosition = "E 0 N 0 U 1.2"
    dblDiameter = 0.055
    dblLength = 2.4
 
    Set ObjShowerVSupport = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, "N 0 U 90")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerVSupport
    Set ObjShowerVSupport = Nothing
    
 ' Insert your code for output 3(Shower Horz. Support)
    strPosition = "E 0 S 0.284 U 2.325"
    dblDiameter = 0.055
    dblLength = 0.568
 
    Set ObjShowerHSupport = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, "S 180 U 0")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerHSupport
    Set ObjShowerHSupport = Nothing
    
 ' Insert your code for output 4(Dish Attach Support)
    strPosition = "E 0 S 0.61 U 2.258"
    dblDiameter = 0.055
    dblLength = 0.05
 
    Set ObjDishAttachSupport = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, "N 0 D -90")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDishAttachSupport
    Set ObjDishAttachSupport = Nothing
    
 ' Insert your code for output 5(Base Support)
    strPosition = "E 0 N 0 U 0.038"
    dblDiameter = 0.229
    dblHeight = 0.076

    Set objCone = CreateCone(m_outputColl, strPosition, "N 0 U 90", dblDiameter, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
  
 ' Insert your code for output 6(Shower Support Bend)
    strPosition = "E 0 S 0.568 U 2.2825"
    dblInRadius = 0.015
    dblOutRadius = 0.07

    Set ObjCTorus = CreateCircularTorus(m_outputColl, strPosition, dblInRadius, dblOutRadius, , "N 0 U 0", "N 0 D -90", 90)
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCTorus
    Set ObjCTorus = Nothing
    '  MsgBox "output 6 torus"
 ' Insert your code for output 7(Shower Head)
    strPosition = "E 0 S 0.61 U 2.159"
    dblDiameter = 0.305
    dblHeight = 0.076
 
    Set ObjDish = CreateDish(m_outputColl, strPosition, "N 0 U 90", dblDiameter, dblHeight, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing
    
' Place Nozzle 1

    Dim NozzlePHFactory As New NozzlePHFactory
    Dim oPipeNozzle As PipeNozzle
    Dim iNozzle As IJDNozzle
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim startPos As IJDPosition
    Dim Axis As IJDVector
    
    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 2, False, m_outputColl.ResourceManager)
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    Set startPos = New DPosition
    startPos.Set 0, 0, 2.4
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set 0, 0, 1
    iDistribPort.SetDirectionVector Axis

    Axis.Set 1, 0, 0
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


' Place Nozzle 2

    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 1, False, m_outputColl.ResourceManager)
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    Set startPos = New DPosition
    startPos.Set 0, 0.06, 0.165
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set 0, 1, 0
    iDistribPort.SetDirectionVector Axis

    Axis.Set 1, 0, 0
    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.06
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPipeNozzle
    Set oPipeNozzle = Nothing
    Set iNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set startPos = Nothing
    Set Axis = Nothing
    

' Place Nozzle 3
    
    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 3, False, m_outputColl.ResourceManager)
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    Set startPos = New DPosition
    startPos.Set -0.06, 0, 0.867
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
    iNozzle.Length = 0.06
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
