VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 5 2004
'   Description:
'       Horizontal Drum with it's Axis along  X-Axis.
'       It has occurences of Outside diameter, Tangent to Tangent, and Insulation Thickness.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjCylinder2 As Object
    Dim ObjDish3 As Object
    Dim ObjDish4 As Object
    
    Dim strPosition As String
    Dim strDirection As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    
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
    strPosition = "E " & parVesselTantoTan / 2 & " N 0 U 0"
    strDirection = "E 90 U 0"
    dblDiameter = parVesselDiameter + 2 * parInsulationThickness
    dblLength = parVesselTantoTan
 
    Set ObjCylinder2 = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
    Set ObjCylinder2 = Nothing
    
 ' Insert your code for output 2(Drum End1)
    Dim dblHeight As Double
 
    strPosition = "E " & parVesselTantoTan & " N 0 U 0"
    strDirection = "E 90 U 0"
    dblDiameter = parVesselDiameter + 2 * parInsulationThickness
    dblHeight = parVesselDiameter / 4 + parInsulationThickness
    Set ObjDish3 = CreateDish(m_outputColl, strPosition, strDirection, dblDiameter, dblHeight, 1, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish3
    Set ObjDish3 = Nothing
    
 ' Insert your code for output 3(Drum End2)
    strPosition = "E 0 N 0 U 0"
    strDirection = "W 270 U 0"
    Set ObjDish4 = CreateDish(m_outputColl, strPosition, strDirection, dblDiameter, dblHeight, 1, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish4
    Set ObjDish4 = Nothing
 
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
