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
'   Author:          MS
'   Creation Date:  Sunday, May 11 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Source for PDB 1046 - ASME 16.48 Spectacle blind
'   Source for PDB 1047 - www.resistoflex.com/pdf/specblind.pdf
'   Source for PDB 1048 - Marine Company Standard for JIS 5K
'   Source for PDB 1049 - http://www.swecofab.com/spectacle/600_female_oct_ring.htm
'   Source for PDB 1050 - www.micromold.com
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006          KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   12.Aug.2008         KKC                 CR-146914  Spectacle blind symbol to be specified by figure 8 inter-axis distance
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInlineSpect As Object
    Dim ObjInterSpectCyl As Object
    Dim ObjDisc As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parFig8InterAxis As Double
    Dim parSpectaclePosition As Double
    Dim parDiscDiameter As Double
    Dim parWebWidth As Double
    Dim parWebThickness As Double
    Dim parRadius As Double
    Dim parCentertoCenter As Double
    Dim parWidth As Double
    Dim parBoltHoleDiameter As Double
    Dim parWebSpacing As Double
    Dim parCenterOffset As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnpoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnpoint = New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    If lPartDataBasis <= 1 Then
        parFacetoFace = arrayOfInputs(2)
        parFig8InterAxis = arrayOfInputs(3)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        If UBound(arrayOfInputs) >= 6 Then
            parDiscDiameter = arrayOfInputs(6)
        End If
    ElseIf lPartDataBasis = 1046 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parWebWidth = arrayOfInputs(8)
        parWebThickness = arrayOfInputs(9)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1047 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parRadius = arrayOfInputs(10)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1048 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parRadius = arrayOfInputs(10)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1049 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parWebWidth = arrayOfInputs(8)
        parWebThickness = arrayOfInputs(9)
        parWebSpacing = arrayOfInputs(11)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1050 Then
        parFacetoFace = arrayOfInputs(2)
        parCenterOffset = arrayOfInputs(12)
        parWidth = arrayOfInputs(13)
        parBoltHoleDiameter = arrayOfInputs(14)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        parCentertoCenter = 2 * parCenterOffset
        parRadius = parWidth / 2
    End If
    iOutput = 0
    If lPartDataBasis <= 1 Then

     ' Insert your code for output 4(Insulated Body)
        Dim Port1Position   As New AutoMath.DPosition
        Port1Position.Set -parFacetoFace / 2, 0, 0
        Dim Port2Position   As New AutoMath.DPosition
        Port2Position.Set parFacetoFace / 2, 0, 0
    
        Dim parInsulationDiameter As Double
        parInsulationDiameter = parFig8InterAxis + 2 * parInsulationThickness
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, parInsulationDiameter, False)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing
        Set Port1Position = Nothing
        Set Port2Position = Nothing
    ElseIf lPartDataBasis = 1046 Or lPartDataBasis = 1047 Or lPartDataBasis = 1048 Or lPartDataBasis = 1049 Or lPartDataBasis = 1050 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        Port1Position.Set -parFacetoFace / 2, 0, 0
        Port2Position.Set parFacetoFace / 2, 0, 0
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, parInsulationDiameter, False)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing
        Set Port1Position = Nothing
        Set Port2Position = Nothing
    End If
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
