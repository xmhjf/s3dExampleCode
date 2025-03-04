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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:   Jun 12, 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   28.Jul.2006     dkl                     Added implementation of PartDataBasis 5
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjEllipsoid As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace  As Double
     
'   Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter   As Double
    Dim pipeDiam1               As Double
    Dim pipeDiam2               As Double
    Dim flangeDiam1             As Double
    Dim flangeDiam2             As Double
    Dim flangeThick1            As Double
    Dim flangeThick2            As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'   parFace1toCenter = arrayOfInputs(2)   'Set only if PartDataBasis=1 or 10
'   parFace2toCenter = arrayOfInputs(3)   'Set only if PartDataBasis=1 or 10
    parInsulationThickness = arrayOfInputs(4)
'   parFacetoFace = arrayOfInputs(6)
    
    iOutput = 0
    
'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis = 5) Then  'Applicable for 5 "Face-to-face dimension basis"
        parFacetoFace = arrayOfInputs(6)
        parFace1toCenter = parFacetoFace / 2  'parFace1toCenter and parFace2toCenter: These dimensions are calculated and used in computation.
        parFace2toCenter = parFacetoFace / 2
    ElseIf (lPartDataBasis = 10) Or (lPartDataBasis <= 1) Then 'Applicable for  1 "Default" & 10 -"Face-to-center dimension basis"
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
    End If
    Set oPipeComponent = Nothing
' Insert your code for output 4(Insulation for Valve Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset, depth

'   Maximum of pipeDiam1, flangeDiam1, pipeDiam2 and flangeDiam2 is used to compute
'   insulation diameter
    parInsulationDiameter = ReturnMax4(pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2) + _
                            2 * parInsulationThickness
                            
    stPoint.Set CenterPos.x - parFace1toCenter, _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x + parFace2toCenter, _
                CenterPos.y, _
                CenterPos.z
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
