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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DAdapterFlange.CAdapterFlange
'   Author:         kkk
'   Creation Date:  Friday, Feb 03 2006
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   01.Aug.2008     MP     CR-146407  Enhance the Adapter symbol to use Seat-to-Seat dimension per JIS G 5527
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim cptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim dblAdapterDia    As Double
    Dim CenterPos        As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    Dim dblGreaterPipeDia As Double
    Dim parFacetoFace As Double
    Dim parSeattoSeat As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    iOutput = 0
    'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    If lPartDataBasis <= 1 Or lPartDataBasis = 20 Or lPartDataBasis = 15 Then
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)
        
    ElseIf lPartDataBasis = 25 Then
        parSeattoSeat = arrayOfInputs(4)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2
        parFacetoFace = parSeattoSeat + depth1 - cptOffset1 + depth2 - cptOffset2
    End If

 ' Insert your code for output 1(Insulated Body)
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2
    
    If CmpDblGreaterthan(pipeDiam1, pipeDiam2) Then
        dblGreaterPipeDia = pipeDiam1
    Else
        dblGreaterPipeDia = pipeDiam2
    End If
    
    If CmpDblGreaterthan(flangeDiam1, flangeDiam2) Then
        dblAdapterDia = flangeDiam1
    Else
      dblAdapterDia = flangeDiam2
    End If
    
    
    If CmpDblEqual(flangeDiam1, 0) Or CmpDblEqual(flangeDiam2, 0) Then
        dblAdapterDia = dblGreaterPipeDia
    End If
    
    Dim parInsulationDiameter As Double
    parInsulationDiameter = dblAdapterDia + parInsulationThickness * 2
    
    Dim endofInsulation       As New AutoMath.DPosition
    endofInsulation.Set parFacetoFace, 0, 0

Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, CenterPos, endofInsulation, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
     
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
