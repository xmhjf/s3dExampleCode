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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'        Gate Valve CL150
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim bonnetDiam           As Double
    Dim discDiam             As Double
    Dim actDiam              As Double
    Dim bodyDiam        As Double
    Dim valveHeight     As Double
    Dim flangeThick     As Double
    Dim hcgs            As Double
    Dim flangeDiam      As Double
    Dim discWidth       As Double
    Dim stemFlWidth     As Double
    Dim stemFlDiam      As Double
    Dim cpDepth         As Double
        
    Dim iOutput     As Double
    Dim ObjCylBodyIns As Object
    Dim ObjVerBodyIns As Object
    
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    discWidth = arrayOfInputs(2)
    valveHeight = arrayOfInputs(3)
    actDiam = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, bodyDiam, flangeThick, flangeDiam, hcgs, cpDepth
    
    If flangeDiam = 0 Then
        discDiam = bodyDiam
        stemFlWidth = bodyDiam * 0.17
        bonnetDiam = discDiam * 1.3
    Else
        discDiam = (bodyDiam + flangeDiam) / 2.2
        stemFlWidth = flangeThick
        bonnetDiam = discDiam
    End If
 ' Insert your code for output 13(Valve Cylindrical Body Insulation)
    Dim InsulationDia As Double
    InsulationDia = bodyDiam + parInsulationThickness * 2
    If flangeDiam > bodyDiam Then InsulationDia = flangeDiam + parInsulationThickness * 2
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -(discWidth / 2), 0, 0
    enPoint.Set (discWidth / 2), 0, 0

     Set ObjCylBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBodyIns
    Set ObjCylBodyIns = Nothing
    
    
 ' Insert your code for output 14(Valve Vertical Body Insulation)
    stPoint.Set -(0.75 * (bonnetDiam / 2)) * 1.2 - parInsulationThickness, _
                    -(1.25 * (bonnetDiam / 2)) - parInsulationThickness, _
                    -(1.25 * (bonnetDiam / 2)) * 1.2 - parInsulationThickness
    enPoint.Set (0.75 * (bonnetDiam / 2)) * 1.2 + parInsulationThickness, _
                        (1.5 * bonnetDiam + 2 * stemFlWidth + 0.002 + 1.25 * (bonnetDiam / 2)) + parInsulationThickness, _
                        (1.25 * (bonnetDiam / 2)) * 1.2 + parInsulationThickness

    Set ObjVerBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerBodyIns
    Set ObjVerBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

