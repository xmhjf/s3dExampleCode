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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Thursday, Jan 25 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.Jun.2007     MA   TR-122032 Modified the computation of the Insulation diameter.
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
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim depth           As Double
    Dim pipeDiam1        As Double
    Dim flangeDiam1      As Double
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parUnionDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parUnionDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)
    Dim stPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                                    flangeDiam1, cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                                    flangeThick, flangeDiam2, cptOffset, depth
    
    'Get the greatest of the Pipe dia Port 1, Flange dia Port 1, Pipe dia Port 2,
    'Flange dia Port 2 and Union Diameter.
    Dim dCylDia1 As Double
    Dim dCylDia2 As Double
    Dim dCylinderInsulationDia As Double
    
    If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
        dCylDia1 = flangeDiam1
    Else
        dCylDia1 = pipeDiam1
    End If
    
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dCylDia2 = flangeDiam2
    Else
        dCylDia2 = pipeDiam2
    End If
    
    If CmpDblGreaterthan(dCylDia1, dCylDia2) Then
        dCylinderInsulationDia = dCylDia1
    Else
        dCylinderInsulationDia = dCylDia2
    End If
    
    If CmpDblGreaterthan(parUnionDiameter, dCylinderInsulationDia) Then _
                                        dCylinderInsulationDia = parUnionDiameter
   
    Dim parInsulationDiameter As Double
    parInsulationDiameter = dCylinderInsulationDia + 2 * parInsulationThickness
    
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


