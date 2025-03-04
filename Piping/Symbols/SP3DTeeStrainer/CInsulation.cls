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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This implements the Insulation aspect of the symbol created for T Strainer to support multiple dimensional bases
'   as per CR-113397
'   The symbol supports the following Part data basis value
'   1.)Part data basis value 85:-'T Strainer, specified by face-to-center dimension',
'      the T strainer symbol will use the Face-to-Center dimension for port 1, port 2,
'      and port 3
'   2.)Part data basis value 90:-T Strainer, specified by Face-to-Face Dimension with Face-to-Center Dimension for Branch
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------     ---      ------------------
'   27.Aug.2007     VRG      TR-124959  Insulation for ports was ensured for different end preparations
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
    Dim ObjInsBranch As Object
    Dim ObjInsStrainer As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsPort3 As Object
    
    Dim parFacetoCenter As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parFacetoFace = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    iOutput = 0
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 85 Then
        parFacetoCenter = arrayOfInputs(2)
    ElseIf lPartDataBasis = 90 Then
        parFacetoCenter = arrayOfInputs(2)
        parFacetoFace = arrayOfInputs(3)
    Else: GoTo ErrorLabel
    End If
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dInsulationDia As Double

' Insert your code for output 4(Insulation for Strainer body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If (lPartDataBasis = 90) Then
        oStPoint.Set (-parFacetoFace / 2) + flangeThick, 0, 0
        oEnPoint.Set (parFacetoFace / 2) - flangeThick, 0, 0
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 85) Then
        oStPoint.Set -parFacetoCenter + flangeThick, 0, 0
        oEnPoint.Set parFacetoCenter - flangeThick, 0, 0
    End If
    
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsStrainer = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsStrainer
    Set ObjInsStrainer = Nothing

 ' Insert your code for output 5(Insulation for Branch)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, parFacetoCenter - flangeThick, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranch
    Set ObjInsBranch = Nothing

 ' Insert your code for output 6(Insulation for Port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    If (lPartDataBasis = 90) Then
        oStPoint.Set (-parFacetoFace / 2), 0, 0
        oEnPoint.Set (-parFacetoFace / 2) + flangeThick + parInsulationThickness, 0, 0
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 85) Then
        oStPoint.Set -parFacetoCenter, 0, 0
        oEnPoint.Set -parFacetoCenter + flangeThick + parInsulationThickness, 0, 0
    End If
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing

 ' Insert your code for output 7(Insulation for Port 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    If (lPartDataBasis = 90) Then
        oStPoint.Set (parFacetoFace / 2), 0, 0
        oEnPoint.Set (parFacetoFace / 2) - flangeThick - parInsulationThickness, 0, 0
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 85) Then
        oStPoint.Set parFacetoCenter, 0, 0
        oEnPoint.Set parFacetoCenter - flangeThick - parInsulationThickness, 0, 0
    End If
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 8(Insulation for Port 3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    oStPoint.Set 0, parFacetoCenter, 0
    oEnPoint.Set 0, parFacetoCenter - flangeThick - parInsulationThickness, 0
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort3
    Set ObjInsPort3 = Nothing

    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
