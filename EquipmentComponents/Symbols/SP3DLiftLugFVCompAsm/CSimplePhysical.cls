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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:11:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Feb 25 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    Dim ObjBox4 As Object
    Dim objPoint As Object
    Dim objPlane As Object
    
    Dim strPosition As String
    Dim dblXLength As Double
    Dim dblYLength As Double
    Dim dblZLength As Double
        
    Dim parVesselDiameter As Double
    Dim parVesselTantoTan As Double
    
    Dim myOri As Orientation
    Set myOri = New Orientation

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 1(Lifting Lug)
    strPosition = "E " & parVesselDiameter / 2 + (0.01 * parVesselDiameter) / 2 & " N 0 U " & parVesselTantoTan + (0.15 * parVesselDiameter) / 2
    dblXLength = 0.01 * parVesselDiameter
    dblYLength = 0.1 * parVesselDiameter
    dblZLength = 0.15 * parVesselDiameter
    myOri.RadianInput = False
    myOri.ResetDefaultAxis
 
    Set ObjBox1 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
' Insert your code for output 2(Lifting Lug)
    strPosition = "E " & parVesselDiameter / -2 - (0.01 * parVesselDiameter) / 2 & " N 0 U " & parVesselTantoTan + (0.15 * parVesselDiameter) / 2
    dblXLength = 0.01 * parVesselDiameter
    dblYLength = 0.1 * parVesselDiameter
    dblZLength = 0.15 * parVesselDiameter

    myOri.ResetDefaultAxis

    Set ObjBox2 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing

' Insert your code for output 3(Lifting Lug)
    strPosition = "E 0 N " & parVesselDiameter / 2 + (0.01 * parVesselDiameter) / 2 & " U " & parVesselTantoTan + (0.15 * parVesselDiameter) / 2
    dblXLength = 0.1 * parVesselDiameter
    dblYLength = 0.01 * parVesselDiameter
    dblZLength = 0.15 * parVesselDiameter

    myOri.ResetDefaultAxis

    Set ObjBox3 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing

' Insert your code for output 4(Lifting Lug)
    strPosition = "E 0 N " & parVesselDiameter / -2 - (0.01 * parVesselDiameter) / 2 & " U " & parVesselTantoTan + (0.15 * parVesselDiameter) / 2
    dblXLength = 0.1 * parVesselDiameter
    dblYLength = 0.01 * parVesselDiameter
    dblZLength = 0.15 * parVesselDiameter

    myOri.ResetDefaultAxis

    Set ObjBox4 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox4
    Set ObjBox4 = Nothing
    Set myOri = Nothing
    
'Insert your code for oputput 5 (Lifting Lug Point)
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
'Insert your code for oputput 5 (Lifting Lug Point)
    strPosition = "E 0 N 0 U 0"
    dblXLength = parVesselDiameter / 2
    dblYLength = dblXLength
    Set objPlane = CreateDefaultPlane(m_outputColl, strPosition, dblXLength, dblYLength, "U")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
