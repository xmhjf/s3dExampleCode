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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 24 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.MAY.2008     MA     CR-141557    Implemented part data basis for the value 13.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parFace2toCenter = arrayOfInputs(5)
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                     sptOffset, depth
                                                                            
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2
    
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -10): FacetoCenter (Default case)
    '   (Part data Basis value -13): Face1toCenter and Face2toCenter
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 10  'FacetoCenter
        parFacetoCenter = arrayOfInputs(2)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13   'Face1toCenter and Face2toCenter
        parFace1toCenter = arrayOfInputs(4)
        parFace2toCenter = arrayOfInputs(5)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    'Insert your code for output 3(Insulation Cylinder)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dInsulationDiameter As Double
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set 0, 0, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    
    'Insert your code for output 3(Insulation Cylinder)
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
    Else
        dInsulationDiameter = pipeDiam2 + parInsulationThickness * 2
    End If
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, dFace2toCenter, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
