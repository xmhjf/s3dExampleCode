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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:05 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Friday, May 16, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedDish As Object
    
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
    
 ' Insert your code for output 1(Insulated Drum Body)
    strPosition = "E 0 N 0 U " & parVesselTantoTan / 2
    strDirection = "N 0 U 90"
    dblDiameter = parVesselDiameter + parInsulationThickness * 2
    dblLength = parVesselTantoTan + parInsulationThickness
 
    Set ObjInsulatedBody = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
' Insert your code for output 2(Insulated Drum End1)
    strPosition = "E 0 N 0 U " & parVesselTantoTan + parInsulationThickness * 0.5
    strDirection = "N 0 U 90"
    dblDiameter = parVesselDiameter + 2 * parInsulationThickness
    
    Set ObjInsulatedDish = Create2to1EllipticalDish(m_outputColl, strPosition, strDirection, dblDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedDish
    Set ObjInsulatedDish = Nothing
    
' Insert your code for output 3(Insulated Drum End2)
    strPosition = "E 0 N 0 D " & parInsulationThickness * 0.5
    strDirection = "N 0 D -90"
    dblDiameter = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedDish = Create2to1EllipticalDish(m_outputColl, strPosition, strDirection, dblDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedDish
    Set ObjInsulatedDish = Nothing
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

