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
'   Copyright (c) 2004 Intergraph
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, May 4 2004
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'       HVAC Square Throat Elbow
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.May.2004     svsmylav        Insulation aspect is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Private Sub Class_Initialize()
'
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    
    Dim ObjHVACNoz1Ins As Object
    Dim ObjHVACNoz2Ins As Object
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parThroat1 As Double
    Dim parThroat2 As Double
    Dim parPlaneOfTurn As Long
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parThroat1 = arrayOfInputs(4)
    parThroat2 = arrayOfInputs(5)
    parPlaneOfTurn = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 3(Nozzle 1 insulation)
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    If parPlaneOfTurn = 1 Then
        stPoint.Set -(parThroat1 + parWidth / 2), _
                    -(parWidth / 2 + parInsulationThickness), _
                    -(parDepth / 2 + parInsulationThickness)
        enPoint.Set parWidth / 2 + parInsulationThickness, _
                    parWidth / 2 + parInsulationThickness, _
                    parDepth / 2 + parInsulationThickness
    Else
        stPoint.Set -(parThroat1 + parDepth / 2), _
                    -(parDepth / 2 + parInsulationThickness), _
                    -(parWidth / 2 + parInsulationThickness)
        enPoint.Set parDepth / 2 + parInsulationThickness, _
                    parDepth / 2 + parInsulationThickness, _
                    parWidth / 2 + parInsulationThickness
    End If

    Set ObjHVACNoz1Ins = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHVACNoz1Ins
    Set ObjHVACNoz1Ins = Nothing

 ' Insert your code for output 4 (Nozzle 2 insulation)
    If parPlaneOfTurn = 1 Then
        stPoint.Set -(parWidth / 2 + parInsulationThickness), _
                    (parThroat2 + parWidth / 2), _
                    -(parDepth / 2 + parInsulationThickness)
        enPoint.Set (parWidth / 2 + parInsulationThickness), _
                     parWidth / 2 + parInsulationThickness, _
                    (parDepth / 2 + parInsulationThickness)
    Else
        stPoint.Set -(parDepth / 2 + parInsulationThickness), _
                    (parThroat2 + parDepth / 2), _
                    -(parWidth / 2 + parInsulationThickness)
        enPoint.Set (parDepth / 2 + parInsulationThickness), _
                     parDepth / 2 + parInsulationThickness, _
                    (parWidth / 2 + parInsulationThickness)
    End If

    Set ObjHVACNoz2Ins = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHVACNoz2Ins
    Set ObjHVACNoz2Ins = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

