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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Jan 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub
Private Function ReturnMax2(A#, B#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    ReturnMax2 = MaxValue
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
    Dim ObjBodyBall As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjInsulatedCylinder1 As Object
    Dim ObjInsulatedCylinder2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    
'   Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationDiameter   As Double
    Dim extendInsulationFlag    As Double
    Dim pipeDiam1               As Double
    Dim pipeDiam2               As Double
    Dim flangeDiam1             As Double
    Dim flangeDiam2             As Double

    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    
    iOutput = 0

' Insert your code for output 4 (Insulation for Cylinder1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    parInsulationDiameter = ReturnMax2(pipeDiam1, flangeDiam1) + 2 * parInsulationThickness
    If ReturnMax2(pipeDiam1, flangeDiam1) > ReturnMax2(pipeDiam2, flangeDiam2) Then
        extendInsulationFlag = 1
    Else
        extendInsulationFlag = 2
    End If
    stPoint.Set CenterPos.x - parFace1toCenter, _
                CenterPos.y, _
                CenterPos.z
    If extendInsulationFlag <> 1 Then
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Else
'       Extend the insulation along positive X axis by Ball radius plus parInsulationThickness
'       (Ball diameter is considered to be half of the face to center distance)
        enPoint.Set CenterPos.x + ReturnMax2(parFace1toCenter, parFace2toCenter) / 4 + parInsulationThickness, _
                    CenterPos.y, _
                    CenterPos.z
    End If
    Set ObjInsulatedCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder1
    Set ObjInsulatedCylinder1 = Nothing
    
' Insert your code for output 5(Insulation for Cylinder2)
    parInsulationDiameter = ReturnMax2(pipeDiam2, flangeDiam2) + 2 * parInsulationThickness
    stPoint.Set CenterPos.x, _
                CenterPos.y + parFace2toCenter, _
                CenterPos.z
    If extendInsulationFlag <> 2 Then
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Else
'       Extend the insulation along negative Y axis by Ball radius plus parInsulationThickness
'       (Ball diameter is considered to be half of the face to center distance)
        enPoint.Set CenterPos.x, _
                    CenterPos.y - (ReturnMax2(parFace1toCenter, parFace2toCenter) / 4 _
                                - parInsulationThickness), _
                    CenterPos.z
    End If
    Set ObjInsulatedCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder2
    Set ObjInsulatedCylinder2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
