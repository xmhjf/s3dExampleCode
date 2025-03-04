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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:19 AM"
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
    Dim ObjCylinder1 As Object
    Dim ObjDish1 As Object
    Dim ObjDish2 As Object
    Dim objPoint As Object
    Dim objPlane As Object
    
    Dim strPosition As String
    Dim strDirection As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    Dim dblXLength As Double
    Dim dblYLength As Double
    
    Dim parVesselDiameter As Double
    Dim parVesselTantoTan As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
       
 ' Insert your code for output 1(Drum Body)
    strPosition = "E 0 N " & parVesselTantoTan / 2 & " U 0"
    strDirection = "N 0 U 0"
    dblDiameter = parVesselDiameter
    dblLength = parVesselTantoTan
 
    Set ObjCylinder1 = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
    
 ' Insert your code for output 2(Drum End1)
    strPosition = "E 0 N " & parVesselTantoTan & " U 0"
    strDirection = "N 0 U 0"
    dblDiameter = parVesselDiameter
    Set ObjDish1 = Create2to1EllipticalDish(m_outputColl, strPosition, strDirection, dblDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish1
    Set ObjDish1 = Nothing
    
 ' Insert your code for output 3(Drum End2)
    strPosition = "E 0 N 0 U 0"
    strDirection = "S 180 U 0"
    dblDiameter = parVesselDiameter
    Set ObjDish2 = Create2to1EllipticalDish(m_outputColl, strPosition, strDirection, dblDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish2
    Set ObjDish2 = Nothing
    
' Insert your code for output 4 (Point)
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    Set CenterPos = Nothing
    
' Insert your code for output 5(Default Plane)
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
