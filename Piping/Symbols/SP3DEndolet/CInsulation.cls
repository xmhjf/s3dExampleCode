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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          Babu Govindarajan
'   Creation Date:  Monday, July 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-81
'   SN=F179 and MC=FOLHC. Symbol consist of Physical and Insulation aspects. Physical aspect is
'   made up of 1.Diverging Tapering Cone, 2.Tapering Cone.
'   Insulation Aspect consist of Insulation for Diverging cone, Converging cone and for Nozzle2
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulatedEndolet As Object
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0
    
 ' Insert your code for output 3(Endolet Insulation)
    Dim InsulationDiameter As Double
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
 'In case port 2 is Bevel end (butt welded); i.e. flangeDiam2 = 0

    If flangeDiam2 = 0 Then
        InsulationDiameter = pipeDiam2 * 1.65 + parInsulationThickness * 2
    Else
        InsulationDiameter = flangeDiam2 * 1.25 + parInsulationThickness * 2
    End If
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition

    startPoint.Set 0, 0, 0
    endPoint.Set parFacetoFace, 0, 0
    Set ObjInsulatedEndolet = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedEndolet
    Set ObjInsulatedEndolet = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
