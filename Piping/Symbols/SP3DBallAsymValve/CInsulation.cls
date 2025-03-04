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
'   Author:         RAH
'   Creation Date:  Oct 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:

'   dd.mmm.yyyy     who             change description
'   -----------     -----           ------------------
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'   08.SEP.2006     KKC             DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG             TR-124959  Insulation for ports was ensured for different end preparations
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
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim iOutput     As Integer
    Dim ObjBallInsulation As Object
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    Dim InsulationDia As Double
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    iOutput = 0
 
' Insert your code for output 6(As a Cylinder)
    Dim InsulationDia1 As Double
    Dim InsulationDia2 As Double
    Dim StartInsulationPos       As New AutoMath.DPosition
    Dim EndInsulationPos         As New AutoMath.DPosition
    StartInsulationPos.Set -parFace1toCenter, 0, 0
    EndInsulationPos.Set parFace2toCenter, 0, 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDia1 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDia1 = pipeDiam + parInsulationThickness * 2
    End If
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDia2 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDia2 = pipeDiam + parInsulationThickness * 2
    End If
    If InsulationDia1 > InsulationDia2 Then
        InsulationDia = InsulationDia1
    Else
        InsulationDia = InsulationDia2
    End If
    Set ObjBallInsulation = PlaceCylinder(m_OutputColl, StartInsulationPos, EndInsulationPos, _
                                    InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBallInsulation
    Set ObjBallInsulation = Nothing
    Set StartInsulationPos = Nothing
    Set EndInsulationPos = Nothing
 
 
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
