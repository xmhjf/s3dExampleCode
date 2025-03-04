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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Jun 9 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is prepared based on PDS symbol MC=TKDR, SN=V62.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.Feb.2008     MA   CR-135752 Implemented Part Data Basis for the value 13.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

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
    Dim CenterPos       As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjInsulatedCylinder1 As Object
    Dim ObjInsulatedCylinder2 As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationDiameter   As Double
    Dim parHandwheelAngle   As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
'    parFace1toCenter = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)
    
    iOutput = 0

    'Insert your code for output 3(Insulation for Cylinder1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    'This symbol is based on the following Part data basis values that govern its geometry
    '(Part data Basis value -10): Face-to-center dimension basis
    '(Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
    
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 10
        parFacetoCenter = arrayOfInputs(2)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parFace1toCenter = arrayOfInputs(5)
        parFace2toCenter = arrayOfInputs(6)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    'Anti-Clock wise angle measurement is taken to be positive
    Dim dValveAng As Double
    dValveAng = 5 * PI / 4

    stPoint.Set CenterPos.x + dFace1toCenter * Cos(dValveAng), _
                CenterPos.y + dFace1toCenter * Sin(dValveAng), _
                CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    
    Set ObjInsulatedCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                    parInsulationDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder1
    Set ObjInsulatedCylinder1 = Nothing
    
    'Insert your code for output 5(Insulation for Cylinder2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    stPoint.Set CenterPos.x + dFace2toCenter, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    
    Set ObjInsulatedCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                    parInsulationDiameter, True)
    
    'Set the output
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
