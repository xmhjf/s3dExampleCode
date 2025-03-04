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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Friday, Aug 06 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  26.Mar.2008     RRK    CR-83045:Implemented part data basis value-'Lap joint flange, backing flange only'(931)
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
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim flangeDiam1     As Double
    Dim cptOffset       As Double
    Dim depth           As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    
    Dim iOutput     As Double
    Dim objFlangeIns As Object
    Dim objStubEndIns As Object
    
    Dim parStubLength As Double
    Dim parLapThickness As Double
    Dim parIsPipeFlared As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    Dim parHubDiameter As Double
    
    Dim dLapThickness As Double
    Dim dPort1Location As Double
    Dim dPort2Location As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '    parStubLength = arrayOfInputs(2)
    '    parLapThickness = arrayOfInputs(3)
    parIsPipeFlared = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    '    parFacetoFace = arrayOfInputs(6)
    '    parHubDiameter = arrayOfInputs(7)

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset, depth, False
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2

    Select Case lPartDataBasis
    Case Is <= 1, 20, 15, 930
        parStubLength = arrayOfInputs(2)
        parLapThickness = arrayOfInputs(3)
        dLapThickness = parLapThickness
    Case 931
        parFacetoFace = arrayOfInputs(6)
        parHubDiameter = arrayOfInputs(7)
        'Retrieving the lap thickness value from bolted end generic data
        Call oPipeComponent.GetLapThickness(1, dLapThickness)
    End Select
    
    iOutput = 0
    
' Create Output 1 for Lap Flange Cylinder Insulation
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim dInsulationDia As Double
     
    If lPartDataBasis = 931 Then
        stPoint.Set dLapThickness, 0, 0
        enPoint.Set flangeThick1 + dLapThickness + parInsulationThickness, 0, 0
        
        'Check to see that insulation doesn't exceed face to face
        If CmpDblGreaterthan(flangeThick1 + parInsulationThickness, parFacetoFace) Then enPoint.x = dLapThickness + parFacetoFace
    Else
        stPoint.Set 0, 0, 0
        enPoint.Set flangeThick1 + dLapThickness + parInsulationThickness, 0, 0
    End If
    
    dInsulationDia = flangeDiam1 + 2 * parInsulationThickness
    Set objFlangeIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeIns
    Set objFlangeIns = Nothing

' Create Output 2 for Stub End Cylinder Insulation
    ' Create Output 2 for insulation of stub end or conical hub of backing flange based on part data basis
    If lPartDataBasis = 931 Then 'Creating conical hub
        Dim ObjInsConicalHub As Object
        Dim dHubInsulationDia As Double
        
        stPoint.Set dLapThickness + flangeThick1, 0, 0
        enPoint.Set dLapThickness + parFacetoFace, 0, 0
        
        Dim dHubDiameter As Double
        dHubDiameter = parHubDiameter
        
        'Considering hub diameter as optional parameter, assigning a value equal to 1.2 times the pipe diameter of non-bolted end
        If CmpDblEqual(dHubDiameter, 0) Then dHubDiameter = 1.2 * pipeDiam2
        
        'Check to see that hub diameter with the above assumption  doesn't exceed flange diameter
        If CmpDblGreaterthanOrEqualTo(dHubDiameter, flangeDiam1) Then dHubDiameter = (flangeDiam1 + pipeDiam2) / 2
        
        dHubInsulationDia = dHubDiameter + 2 * parInsulationThickness
        dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
        
        Set ObjInsConicalHub = PlaceCone(m_OutputColl, stPoint, enPoint, dHubInsulationDia / 2, dInsulationDia / 2, True)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsConicalHub
        Set ObjInsConicalHub = Nothing
    Else
        
        dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
            
        If parIsPipeFlared <> 0 Then ''For a Flared pipe the stublength would be zero.
            parStubLength = 0.000001
            dInsulationDia = pipeDiam2
        End If
        
        stPoint.Set 0, 0, 0
        enPoint.Set parStubLength, 0, 0

        Set objStubEndIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStubEndIns
        Set objStubEndIns = Nothing
    End If
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
