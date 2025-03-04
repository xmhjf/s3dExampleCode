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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:20:02 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          CYW
'   Creation Date:  Tuesday, Dec 10 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjBase As Object
    Dim ObjStep As Object
    Dim ObjBack As Object
    Dim ObjNeck As Object
    Dim ObjHead As Object
    Dim ObjMeter As Object
    Dim ObjArrow As Object

    Dim oriOrientation As Orientation
    Set oriOrientation = New Orientation
    
' Inputs
'    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0

'******
' Insert your code for output 1(Base)
'******
    Set ObjBase = CreateBox(m_outputColl, "W 0.3175 S 0.508 U 0.127", oriOrientation, 0.635, 1.016, 0.254, True)
    
'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBase
    Set ObjBase = Nothing
        
'******
' Insert your code for output 2(Step)
'******
    Set ObjStep = CreateBox(m_outputColl, "W 0.32385 S 0.5842 U 0.2794", oriOrientation, 0.5842, 0.8128, 0.0254, True)
    
'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjStep
    Set ObjStep = Nothing
        
'******
' Insert your code for output 3(Back)
'******
    Set ObjBack = CreateBox(m_outputColl, "W 0.3175 S 0.0762 U 0.7112", oriOrientation, 0.5588, 0.1524, 0.9144, True)

'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBack
    Set ObjBack = Nothing

'******
' Insert your code for output 4(Neck)
'******
    Set ObjNeck = CreateBox(m_outputColl, "W 0.3175 S 0.0762 U 1.22079", oriOrientation, 0.381, 0.127, 0.104775, True)

'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjNeck
    Set ObjNeck = Nothing

'******
' Insert your code for output 5(Head)
'******
    Set ObjHead = CreateCylinder(m_outputColl, "W 0.3175 S 0.0762 U 1.524", 0.635, 0.1524, "S 180 U 0")

'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjHead
    Set ObjHead = Nothing
        
'******
' Insert your code for output 6(Meter)
'******
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutX = 90
    oriOrientation.RotationAboutY = 45
    oriOrientation.ApplyRotations
    Set ObjMeter = CreateBox(m_outputColl, "W 0.33496 S 0.15557 U 1.50654", oriOrientation, 0.0254, 0.3048, 0.00635, True)

'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjMeter
    Set ObjMeter = Nothing
        
'******
' Insert your code for output 7(Arrow)
'******
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 45
    oriOrientation.ApplyRotations
    Set ObjArrow = CreatePyramid(m_outputColl, "W 0.20955 S 0.15875 U 1.63195", oriOrientation, 0.0508, 0.0127, 0#, 0.0127, 0#, 0#, 0.0762)


'--- Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjArrow
    Set ObjArrow = Nothing
        
        
        
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
