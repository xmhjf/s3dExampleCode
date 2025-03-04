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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:14 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, Apr 3 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol is created using an insulation for the vessel.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports As Double
    Dim parVesselCenterHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportRadialLocation As Double
    Dim parSupportHeight As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)     'P2
    parSupportAngularLocation = arrayOfInputs(3) 'P3
    parNumberOfSupports = arrayOfInputs(4)   'P4
    parVesselCenterHeight = arrayOfInputs(5) 'P5
    parSupportLength = arrayOfInputs(6)      'P6
    parSupportThickness = arrayOfInputs(7)   'P7
    parSupportRadialLocation = arrayOfInputs(8) 'P8
    parSupportHeight = arrayOfInputs(9)      'P9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

' Insert your code for output 1(InsulatedBody)
    parInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    Dim centPoint As New AutoMath.DPosition
    centPoint.Set CenterPos.x, _
                  CenterPos.y, _
                  CenterPos.z + parVesselCenterHeight
    Set ObjInsulatedBody = PlaceSphere(m_OutputColl, centPoint, parInsulationDiameter)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
