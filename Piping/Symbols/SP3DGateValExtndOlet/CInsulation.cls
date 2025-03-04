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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         BG
'   Creation Date:  Monday, Jun 17 2002
'   Description:
'   The Symbol details were taken from PDS PDS Piping Component Data Reference Manual
'   at Page No D - 8 and SN=V2A. The Symbol consist of Physical and Insulation aspects
'   Physical aspect is made up of three cones and two Nozzles.Insulation aspect consist of
'   Simple Cylinder between flange of 2nd Nozzle and to the point where body extension ends
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'   09.Sep.2005     svsmylav        TR-83135: Modified the Y-coordinates of points in the
'                                           symbol by adding 'parValveCentertoHeaderCenter'.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CInsulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulationCylin As Object
    
    Dim parValveCentertoHeaderCenter As Double
    Dim parFacetoCenter As Double
    Dim parOletLength As Double
    Dim parInsulationThickness As Double

    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveCentertoHeaderCenter = arrayOfInputs(2)
    parFacetoCenter = arrayOfInputs(3)
    parOletLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 4(InsulationCylin)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim InsulationDia As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, parValveCentertoHeaderCenter + parFacetoCenter, 0
    endPoint.Set 0, parValveCentertoHeaderCenter - parFacetoCenter - parOletLength, 0
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDia = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDia = pipeDiam + parInsulationThickness * 2
    End If
    Set ObjInsulationCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationCylin
    Set ObjInsulationCylin = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
