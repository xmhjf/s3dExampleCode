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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I120AZ Page No. 363 of PDS Piping Component Data Reference Guide pcd.pdf.
'   Change History:
'   dd.mmm.yyyy     who           change description
'   -----------     -----         ------------------
'   08.SEP.2006     KKC           DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.Aug.2007     VRG           TR-124960 Insulation diameter was computed based on End preparation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Integer
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedNozzzle As Object
    Dim parFacetoEnd As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoEnd = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)
    Dim PlugEndPosition   As New AutoMath.DPosition
    PlugEndPosition.Set parFacetoEnd + parInsulationThickness, 0, 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parInsulationDiameter As Double
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, CenterPos, PlugEndPosition, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set PlugEndPosition = Nothing
    
  
 ' Insert your code for output 4(Insulated Nozzle)
    Dim InsNozlEndPosition   As New AutoMath.DPosition
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = NEGLIGIBLE_THICKNESS
    End If
    InsNozlEndPosition.Set flangeThick + parInsulationThickness, 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    Set ObjInsulatedNozzzle = PlaceCylinder(m_OutputColl, CenterPos, InsNozlEndPosition, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedNozzzle
    Set ObjInsulatedNozzzle = Nothing
    Set CenterPos = Nothing
    Set InsNozlEndPosition = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
      
End Sub
