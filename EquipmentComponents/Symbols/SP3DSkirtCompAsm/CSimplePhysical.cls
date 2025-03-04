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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:55 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          nka8411
'   Creation Date:  Tuesday, Feb 25 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
 
    Dim iOutput     As Double
    Dim ObjSnout As Object
    Dim ObjPoint As Object
    Dim objPlane1 As Object
    Dim objPlane2 As Object
    
    Dim strPosition As String
    Dim dblTopDiameter As Double
    Dim dblBotDiameter As Double
    Dim dblLength As Double
    Dim dblXLength As Double
    Dim dblYLength As Double
    
    Dim parVesselDiameter As Double
    Dim parVesselTantoTan As Double
    
    Dim myOri As Orientation
    Set myOri = New Orientation

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 1(Drum Support Skirt)
    strPosition = "E 0 N 0 D " & (parVesselTantoTan * 0.5) / 2
    dblTopDiameter = parVesselDiameter
    dblBotDiameter = parVesselDiameter * 1.3
    dblLength = parVesselTantoTan * 0.5
    
    myOri.RadianInput = False
    myOri.ResetDefaultAxis
    myOri.RotationAboutY = -90
    myOri.ApplyRotations
 
    Set ObjSnout = CreateSnout(m_outputColl, strPosition, myOri, 0, dblTopDiameter, dblBotDiameter, dblLength, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSnout
    Set ObjSnout = Nothing

'Insert your code for oputput 2(Lifting Lug Point)
    Set ObjPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPoint
    Set ObjPoint = Nothing
    
'Insert your code for oputput 3(Default Plane)
    strPosition = "E 0 N 0 D " & parVesselTantoTan * 0.5
    dblXLength = parVesselDiameter / 2
    dblYLength = dblXLength
    Set objPlane1 = CreateDefaultPlane(m_outputColl, strPosition, dblXLength, dblYLength, "U")
   
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput "Default_Plane", objPlane1
    Set objPlane1 = Nothing

'Insert your code for oputput 4(Default Plane)
    strPosition = "E 0 N 0 U 0"
    dblXLength = parVesselDiameter / 2
    dblYLength = dblXLength
    Set objPlane2 = CreateDefaultPlane(m_outputColl, strPosition, dblXLength, dblYLength, "D")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput "Plane_Default", objPlane2
    Set objPlane2 = Nothing
    
    Set myOri = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
