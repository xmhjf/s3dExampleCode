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
'   Creation Date:  Monday, Nov 21, 2005
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is Trimmable Generic Cross symbol. (PDS Reducing run and branches cross (MC=XRRB)
'     (SN=F163)Symbol provides a close example to this symbol).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjPort1Insulation As Object
    Dim ObjPort2Insulation As Object
    Dim ObjPort3Insulation As Object
    Dim ObjPort4Insulation As Object
    Dim ObjMainBodyInsul As Object
    Dim ObjTaperConeInsul As Object
    Dim ObjBranchInsulation As Object
    Dim ObjBranchInsulation1 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'A (of PDS SN=F163)
    parFace2toCenter = arrayOfInputs(3) 'B
    parFace3toCenter = arrayOfInputs(4) 'C
    parFace4toCenter = arrayOfInputs(5) 'D
    parAngle = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 6(Insulation for Port 1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dInsulationDia As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjPort1Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort1Insulation
    Set ObjPort1Insulation = Nothing

' Insert your code for output 10(Insulation for the Main Body)
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjMainBodyInsul = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainBodyInsul
    Set ObjMainBodyInsul = Nothing

' Insert your code for output 7(Insulation for Port 2)
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = 0.0001
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set (parFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
    dInsulationDia = pipeDiam2 + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then dInsulationDia = flangeDiam2 + parInsulationThickness * 2
    Set ObjPort2Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort2Insulation
    Set ObjPort2Insulation = Nothing
    
' Insert your code for output 11(Insulation for Taper Cone)
    stPoint.Set 0, 0, 0
    enPoint.Set parFace2toCenter - flangeThick2, 0, 0
    
    Set ObjTaperConeInsul = PlaceCone(m_OutputColl, stPoint, enPoint, _
          pipeDiam * 0.5 + parInsulationThickness, pipeDiam2 * 0.5 + parInsulationThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperConeInsul
    Set ObjTaperConeInsul = Nothing
    
 ' Insert your code for output 8(Insulation for Port 3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
    enPoint.Set (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
        (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjPort3Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Insulation
    Set ObjPort3Insulation = Nothing

' Insert your code for output 12(Insulation for Branch)
    stPoint.Set (parFace3toCenter - flangeThick) * Cos(parAngle), _
        (parFace3toCenter - flangeThick) * Sin(parAngle), 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjBranchInsulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation
    Set ObjBranchInsulation = Nothing

' Insert your code for output 9(Insulation for Port 4)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set parFace4toCenter * Cos(parAngle + PI), parFace4toCenter * Sin(parAngle + PI), 0
    enPoint.Set (parFace4toCenter - flangeThick - parInsulationThickness) * Cos(parAngle + PI), _
        (parFace4toCenter - flangeThick - parInsulationThickness) * Sin(parAngle + PI), 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjPort4Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort4Insulation
    Set ObjPort4Insulation = Nothing
        
 ' Insert your code for output 13(Insulation for Branch 1)
    stPoint.Set (parFace4toCenter - flangeThick) * Cos(parAngle + PI), _
        (parFace4toCenter - flangeThick) * Sin(parAngle + PI), 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjBranchInsulation1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation1
    Set ObjBranchInsulation1 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
