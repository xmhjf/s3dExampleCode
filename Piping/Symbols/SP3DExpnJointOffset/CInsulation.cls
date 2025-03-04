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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 3 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'   16.Feb.2006     svsmylav        RI -28367: Updated the symbol to use Part Data Basis property
'                                           and added new input parameter 'FacetoFace' for Model 'AUD',Metal%20Catalog.pdf,
'                                           of Flexonics, URL http://www.flexonics.ca/products.htm.
'                                           Modified code to use CmpDblGreaterthan for double comparision.
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
    Dim depth           As Double
    Dim iOutput     As Double
    Dim ObjNozzle1Insul As Object
    Dim ObjNoz1ExtnInsul As Object
    Dim ObjJunction1Insul As Object
    Dim ObjSkewPipeInsul As Object
    Dim ObjJunction2Insul As Object
    Dim ObjNoz2ExtnInsul As Object
    Dim ObjNozzle2Insul As Object
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset As Double
    Dim parFacetoFace As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
'    parFacetoFace = arrayOfInputs(6)

    iOutput = 0
    
 'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If (lPartDataBasis <> 5) Then
        'Consider Face-to-center dimension basis
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
    
    ElseIf (lPartDataBasis = 5) Then
        'Face-to-face dimension basis
        parFacetoFace = arrayOfInputs(6)
        'Consider parFace1toCenter and parFace2toCenter to be half of parFacetoFace
        parFace1toCenter = parFacetoFace / 2
        parFace2toCenter = parFacetoFace / 2
    End If

 ' Insert your code for output 6(Insulation for Nozzle1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    InsulationDia = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then InsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjNozzle1Insul = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1Insul
    Set ObjNozzle1Insul = Nothing
    
 ' Insert your code for output 7(Insulation Nozzle1 Extension)
    Dim NozInsStPoint   As New AutoMath.DPosition
    Dim NozInsEnPoint   As New AutoMath.DPosition
    Dim NozInsulationDia As Double
    NozInsStPoint.Set -parFace1toCenter + flangeThick, 0, 0
    NozInsEnPoint.Set -(parFace1toCenter / 3 + parFace1toCenter / 3), 0, 0
    NozInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjNoz1ExtnInsul = PlaceCylinder(m_OutputColl, NozInsStPoint, NozInsEnPoint, NozInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1ExtnInsul
    Set ObjNoz1ExtnInsul = Nothing
    
 ' Insert your code for output 8(Junction 1 Insulation)
    Dim JunInsStPoint   As New AutoMath.DPosition
    Dim JunInsEnPoint   As New AutoMath.DPosition
    Dim JunInsulationDia As Double
    JunInsStPoint.Set -(parFace1toCenter / 3 + parFace1toCenter / 3 + parInsulationThickness), 0, 0
    JunInsEnPoint.Set -(parFace1toCenter / 3 - parInsulationThickness), 0, 0
    JunInsulationDia = pipeDiam * 1.5 + parInsulationThickness * 2
    Set ObjJunction1Insul = PlaceCylinder(m_OutputColl, JunInsStPoint, JunInsEnPoint, JunInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjJunction1Insul
    Set ObjJunction1Insul = Nothing
    
 ' Insert your code for output 9(Insulation for Skew Pipe)
    Dim Angle As Double
    Dim Hypotenous  As Double
    Dim SkewInsStPoint   As New AutoMath.DPosition
    Dim SkewInsEnPoint   As New AutoMath.DPosition
    Dim SkewInsulationDia As Double
    Angle = Atn(parOffset / (parFace1toCenter / 3))
    Hypotenous = 0.5 * pipeDiam * Tan(Angle)
    SkewInsStPoint.x = -(Hypotenous * Cos(Angle) + parFace1toCenter / 3)
    SkewInsStPoint.y = -Hypotenous * Sin(Angle)
    SkewInsStPoint.z = 0
    SkewInsEnPoint.x = (Hypotenous * Cos(Angle) + parFace2toCenter / 3)
    SkewInsEnPoint.y = (parOffset + (parFace2toCenter / parFace1toCenter) * parOffset) + _
                        Hypotenous * Sin(Angle)
    SkewInsEnPoint.z = 0
    SkewInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjSkewPipeInsul = PlaceCylinder(m_OutputColl, SkewInsStPoint, SkewInsEnPoint, SkewInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSkewPipeInsul
    Set ObjSkewPipeInsul = Nothing
    
 ' Insert your code for output 10(Junction 2 Insulation)
    Dim JunAInsStPoint   As New AutoMath.DPosition
    Dim JunAInsEnPoint   As New AutoMath.DPosition
    Dim JunAInsulationDia As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    JunAInsStPoint.Set (parFace2toCenter / 3 - parInsulationThickness), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    JunAInsEnPoint.Set (parFace2toCenter / 3 + parFace2toCenter / 3 + parInsulationThickness), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    JunAInsulationDia = pipeDiam * 1.5 + parInsulationThickness * 2
    Set ObjJunction2Insul = PlaceCylinder(m_OutputColl, JunAInsStPoint, JunAInsEnPoint, JunAInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjJunction2Insul
    Set ObjJunction2Insul = Nothing
    
 ' Insert your code for output 11(Insulation for Nozzle 2 Extn)
    Dim NozAInsStPoint   As New AutoMath.DPosition
    Dim NozAInsEnPoint   As New AutoMath.DPosition
    Dim NozAInsulationDia As Double
    NozAInsStPoint.Set (parFace2toCenter / 3 + parFace2toCenter / 3), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    NozAInsEnPoint.Set (parFace2toCenter - flangeThick), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    NozAInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjNoz2ExtnInsul = PlaceCylinder(m_OutputColl, NozAInsStPoint, NozAInsEnPoint, NozAInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2ExtnInsul
    Set ObjNoz2ExtnInsul = Nothing
    
 ' Insert your code for output 12(Insulation for Nozzle 2)
    Dim Noz2InsStPoint   As New AutoMath.DPosition
    Dim Noz2InsEnPoint   As New AutoMath.DPosition
    Dim Noz2InsulationDia As Double
    Noz2InsStPoint.Set parFace2toCenter, parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    Noz2InsEnPoint.Set (parFace2toCenter - flangeThick - parInsulationThickness), _
                        parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    Noz2InsulationDia = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then Noz2InsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjNozzle2Insul = PlaceCylinder(m_OutputColl, Noz2InsStPoint, Noz2InsEnPoint, Noz2InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2Insul
    Set ObjNozzle2Insul = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
