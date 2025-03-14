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
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Nov 28 2005
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is prepared based on PDS symbol MC=S90YB, S90YRB, SN=F152.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort3 As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'E
    parFace2toCenter = arrayOfInputs(3) 'F
    parFace3toCenter = arrayOfInputs(4) 'G
    parInsulationThickness = arrayOfInputs(5)

    iOutput = 0

' Insert your code for output 1 (Insulated Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset, depth

    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = 0.0001

    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set parFace2toCenter - flangeThick2, 0, 0
    InsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
' Insert your code for output 2 (Insulated Port1)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then InsulationDia = flangeDiam + 2 * parInsulationThickness

    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

' Insert your code for output 3(Insulated Port2)
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set (parFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
    InsulationDia = pipeDiam2 + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then InsulationDia = flangeDiam2 + 2 * parInsulationThickness
    
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing

' Insert your code for output 4(Insulated Branch)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    InsulationDia = pipeDiam + parInsulationThickness * 2
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parFace3toCenter - flangeThick, 0
    
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing
  
' Insert your code for output 5 (Insulated Port3)
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, (parFace3toCenter - flangeThick - parInsulationThickness), 0
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then InsulationDia = flangeDiam + 2 * parInsulationThickness
    
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
