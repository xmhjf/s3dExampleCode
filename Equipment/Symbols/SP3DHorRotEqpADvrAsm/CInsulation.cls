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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:13 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Monday, Apr 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is E405 Horizontal Rotating Equipment and Driver symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjInsulatedEqp As Object
    
    Dim parBaseLength As Double
    Dim parBaseWidth1 As Double
    Dim parBaseWidth2 As Double
    Dim parBaseThickness As Double
    Dim parDistBaseLeftToDischarge As Double
    Dim parDistEqpLeftToDischarge As Double
    Dim parDistEqpRightToDischarge As Double
    Dim parEqpWidth1 As Double
    Dim parEqpWidth2 As Double
    Dim parEqpDepth1 As Double
    Dim parEqpDepth2 As Double
    Dim parRotEqpLength As Double
    Dim parRotEqpDiameter As Double
    Dim parDriverLength As Double
    Dim parDriverWidth1 As Double
    Dim parDriverWidth2 As Double
    Dim parDriverHeight As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseLength = arrayOfInputs(2)
    parBaseWidth1 = arrayOfInputs(3)
    parBaseWidth2 = arrayOfInputs(4)
    parBaseThickness = arrayOfInputs(5)
    parDistBaseLeftToDischarge = arrayOfInputs(6)
    parDistEqpLeftToDischarge = arrayOfInputs(7)
    parDistEqpRightToDischarge = arrayOfInputs(8)
    parEqpWidth1 = arrayOfInputs(9)
    parEqpWidth2 = arrayOfInputs(10)
    parEqpDepth1 = arrayOfInputs(11)
    parEqpDepth2 = arrayOfInputs(12)
    parRotEqpLength = arrayOfInputs(13)
    parRotEqpDiameter = arrayOfInputs(14)
    parDriverLength = arrayOfInputs(15)
    parDriverWidth1 = arrayOfInputs(16)
    parDriverWidth2 = arrayOfInputs(17)
    parDriverHeight = arrayOfInputs(18)
    parInsulationThickness = arrayOfInputs(19)
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0

' Insert your code for output 1(Insulated Equipment)
    
    stPoint.Set parDistBaseLeftToDischarge - parDistEqpLeftToDischarge - parInsulationThickness, _
                parEqpWidth2 + parInsulationThickness, _
                parEqpDepth1 + parInsulationThickness
    enPoint.Set parDistBaseLeftToDischarge + parDistEqpRightToDischarge + parInsulationThickness, _
                -parEqpWidth1 - parInsulationThickness, _
                -parEqpDepth2
    
    Set ObjInsulatedEqp = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedEqp
    Set ObjInsulatedEqp = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
