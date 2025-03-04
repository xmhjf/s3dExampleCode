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
'   Author:         MS
'   Creation Date:  Wednesday, Jan 14, 2004
'   Description:
'   This symbol is created from SP3DReducingCross.CReducingCross, and cache property has been added.
'   Change History:
'   dd.mmm.yyyy       who            change description
'   -----------      -----          ------------------
'   08.SEP.2006       KKC            DI-95670  Replace names with initials in all revision history sheets and symbols
'   06.Aug.2007       VRG            TR-124961 Determination of flange diameter based on flange thickness is avoided
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
  
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Integer
   
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedPort3 As Object
    Dim ObjInsulatedPort4 As Object
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace2toCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 1(Insulated Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    stPoint.Set -(parFacetoCenter - flangeThick), 0, 0
    enPoint.Set parFacetoCenter - flangeThick, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
' Insert your code for output 2(Insulated Port1)
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set -(parFacetoCenter - flangeThick - parInsulationThickness), 0, 0
   
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
         parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

 ' Insert your code for output 3(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set parFacetoCenter, 0, 0
    enPoint.Set (parFacetoCenter - flangeThick - parInsulationThickness), 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
         parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing

  ' Insert your code for output 4(Insulated Branch)
    Dim BranchPoint   As New AutoMath.DPosition
    Dim BranchPoint4   As New AutoMath.DPosition
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    BranchPoint.Set CenterPos.x, parFace2toCenter - flangeThick, CenterPos.z
    BranchPoint4.Set CenterPos.x, -(parFace2toCenter - flangeThick), CenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, BranchPoint4, BranchPoint, parInsulationDiameter, True)
  ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing
  
' Insert your code for output 5 (Insulated Port3)
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set 0, parFace2toCenter, 0
    enPoint.Set 0, (parFace2toCenter - flangeThick - parInsulationThickness), 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
         parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing

' Insert your code for output 5 (Insulated Port4)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, -(parFace2toCenter - flangeThick - parInsulationThickness), 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
         parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    Set ObjInsulatedPort4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort4
    Set ObjInsulatedPort4 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
