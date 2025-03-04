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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:06:09 AM"
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
'    Dim pipeDiam        As Double
'    Dim flangeThick     As Double
'    Dim cptOffset       As Double
'    Dim flangeDiam      As Double
'    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjShowerBase As Object
    Dim ObjDish As Object
    
    Dim strPosition As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    Dim dblHeight As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0

 ' Insert your code for output 1(Shower Base)
 strPosition = "E 0 N 0 U 0.005"
 dblDiameter = 1.829
 dblLength = 0.01
 
 Set ObjShowerBase = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, "N 0 D -90")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerBase
    Set ObjShowerBase = Nothing
    
 ' Insert your code for output 2(Shower Head)
 strPosition = "E 0 N 0 U 2.363"
 dblDiameter = 0.305
 dblHeight = 0.076
 
 Set ObjDish = CreateDish(m_outputColl, strPosition, "N 0 U 90", dblDiameter, dblHeight, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing
    
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
    startPos.Set 0, 0, 2.529
    iLogicalDistPort.SetCenterLocation startPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set Axis = New DVector
    Axis.Set 0, 0, 1
    iDistribPort.SetDirectionVector Axis

    Axis.Set 0, 1, 0
    iDistribPort.SetRadialOrient Axis

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = 0.09
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
