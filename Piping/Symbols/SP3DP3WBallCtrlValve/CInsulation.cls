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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Sep 24 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I14AZ Page No. 358 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private InchToMtr       As Double

Private Sub Class_Initialize()
    InchToMtr = 25.4 / 1000
End Sub
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double

    Dim ObjInsulatedCylinder1 As Object
    Dim ObjInsulatedCylinder2 As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    
'   Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim BallDiameter            As Double
    Dim parInsulationDiameter   As Double
    Dim pipeDiam1               As Double
    Dim pipeDiam2               As Double
    Dim flangeDiam1             As Double
    Dim flangeDiam2             As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
    
' Insert your code for output 5(Insulation for Cylinder1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

'   Maximum of pipeDiam1, flangeDiam1, pipeDiam2 and flangeDiam2 is used to compute
'   insulation diameter
    parInsulationDiameter = ReturnMax4(pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2) + _
                            2 * parInsulationThickness
    stPoint.Set CenterPos.x - parFacetoCenter, _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x + parFace1toCenter, _
                CenterPos.y, _
                CenterPos.z
    Set ObjInsulatedCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder1
    Set ObjInsulatedCylinder1 = Nothing
    
' Insert your code for output 6(Insulation for Cylinder2)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    stPoint.Set CenterPos.x, _
                CenterPos.y + parFace2toCenter, _
                CenterPos.z
'   Extend the insulation along negative Y axis by Ball radius plus parInsulationThickness
 '  Following logic is taken from PDS EDEN
    Dim factor As Double
    If (pipeDiam1 <= 2 * InchToMtr) Then
            factor = 0.5
    Else
            If (pipeDiam1 > 2 * InchToMtr And pipeDiam1 <= 10 * InchToMtr) Then
                    factor = 0.55
            Else
                    If (pipeDiam1 > 10 * InchToMtr) Then
                            factor = 0.6
                    End If
            End If
    End If
    Dim sphere_radius As Double
    sphere_radius = parFacetoCenter * factor * 0.5

    enPoint.Set CenterPos.x, _
                CenterPos.y - sphere_radius - parInsulationThickness, _
                CenterPos.z
                
    If flangeDiam > pipeDiam Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    
    Set ObjInsulatedCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder2
    Set ObjInsulatedCylinder2 = Nothing
    Set CenterPos = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


