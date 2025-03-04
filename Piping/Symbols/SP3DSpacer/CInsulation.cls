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
'   Author:          BG
'   Creation Date:  Monday, Nov 25 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol constructed as per the catalog available at URL http://www.resistoflex.com, all
'   assumptions are based on the catalog provided at the URL location.
'   Physical aspect consist of one  Spacer cylinder and Two nozzles. Insulation aspect consist of
'   a simple cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  29.Jan.2007     dkl  CR 112819 Updated the computation of Spacer diameter in case of non-flanged End preparation.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim raisedFaceOrSocketDiam As Double
    
    Dim iOutput     As Double
    Dim objSpacerIns As Object
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 2(Spacer Insulation)
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Dim dInsulationDia As Double
    Dim dSpacerDia As Double
    
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                    flangeDiam, sptOffset, depth, raisedFaceOrSocketDiam
    
    'If the end is flanged or threaded ie. raised face or socket diameter and flange diameter
    'are non zero then, spacer diameter is the maximum of raised Face Diameter and flange Diameter
     If CmpDblGreaterthan(raisedFaceOrSocketDiam, LINEAR_TOLERANCE) Or _
                            CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
        dSpacerDia = BodyOD(raisedFaceOrSocketDiam, flangeDiam)
    'Else spacer diameter is equal to pipe diameter (End preparation is not flanged).
    Else:  dSpacerDia = pipeDiam
    End If
    
    dInsulationDia = dSpacerDia + 2 * parInsulationThickness
    Set oStPoint = New AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEndPoint.Set parFacetoFace / 2, 0, 0
    
    Set objSpacerIns = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSpacerIns
    Set objSpacerIns = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
