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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          SSP
'   Creation Date:  Friday, Dec 27 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   PDMS Category Name=IVPQ). Physical aspect consist of Relief valve body of Box shape, 3Cylinders
'   in the (-)ive side of Z axis,1 Cylinder in (+)ive side of Z-axis, Nozzle-1 along (+)ive side of X-Axis and
'   Nozzle-2 along (-)ive side of Y-axis.
'   Insulation aspect consist of a single box place covering the whole valve.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
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
    
    Dim iOutput     As Double
    
    Dim ObjInsBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentRadius As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentRadius1 As Double
    Dim parInstrumentHeight4 As Double
    Dim parInstrumentHeight5 As Double
    Dim parInstrumentRadius2 As Double
    Dim parInstrumentHeight6 As Double
    Dim parInstrumentHeight7 As Double
    Dim parInstrumentRadius3 As Double
    Dim parInstrumentHeight8 As Double
    Dim parInstrumentHeight9 As Double
    Dim parCylOffset As Double
    Dim parCylOffset1 As Double
    Dim parInsulationThickness As Double
    
'Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
    
    Dim parPortIndex2 As Long
    Dim parNpd2 As Double
    Dim parEndPreparation2 As Long
    Dim parScheduleThickness2 As Long
    Dim parEndStandard2 As Long
    Dim parPressureRating2 As Long
    Dim parFlowDirection2 As DistribFlow
    Dim parId2 As String
    Dim parNPDUnitType2 As String
    Dim parOperationLength As Double
    Dim parOperationHeight As Double
    Dim parOperationWidth As Double
    Dim parOperationWidth1 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)             'D2
    parFace2toCenter = arrayOfInputs(3)             'D3
    parInstrumentHeight = arrayOfInputs(4)          'D4
    parInstrumentHeight1 = arrayOfInputs(5)         'D5
    parInstrumentLength = arrayOfInputs(6)          'D6
    parInstrumentLength1 = arrayOfInputs(7)         'D7
    parInstrumentWidth = arrayOfInputs(8)           'D8
    parInstrumentWidth1 = arrayOfInputs(9)          'D9
    parInstrumentRadius = arrayOfInputs(10)         'D10
    parInstrumentHeight2 = arrayOfInputs(11)        'D11
    parInstrumentHeight3 = arrayOfInputs(12)        'D12
    parInstrumentRadius1 = arrayOfInputs(13)        'D13
    parInstrumentHeight4 = arrayOfInputs(14)        'D14
    parInstrumentHeight5 = arrayOfInputs(15)        'D15
    parInstrumentRadius2 = arrayOfInputs(16)        'D16
    parInstrumentHeight6 = arrayOfInputs(17)        'D17
    parInstrumentHeight7 = arrayOfInputs(18)        'D18
    parInstrumentRadius3 = arrayOfInputs(19)        'D19
    parInstrumentHeight8 = arrayOfInputs(20)        'D20
    parInstrumentHeight9 = arrayOfInputs(21)        'D21
    parCylOffset = arrayOfInputs(22)        'D22
    parCylOffset1 = arrayOfInputs(23)        'D23
    parInsulationThickness = arrayOfInputs(24)
    parNPD = arrayOfInputs(25)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(26)
    parScheduleThickness = arrayOfInputs(27)
    parEndStandard = arrayOfInputs(28)
    parPressureRating = arrayOfInputs(29)
    parFlowDirection = arrayOfInputs(30)
    parPortIndex1 = arrayOfInputs(31)
    parNpd1 = arrayOfInputs(32)
    parEndPreparation1 = arrayOfInputs(33)
    parScheduleThickness1 = arrayOfInputs(34)
    parEndStandard1 = arrayOfInputs(35)
    parPressureRating1 = arrayOfInputs(36)
    parFlowDirection1 = arrayOfInputs(37)
    parPortIndex2 = arrayOfInputs(38)
    parNpd2 = arrayOfInputs(39)
    parEndPreparation2 = arrayOfInputs(40)
    parScheduleThickness2 = arrayOfInputs(41)
    parEndStandard2 = arrayOfInputs(42)
    parPressureRating2 = arrayOfInputs(43)
    parFlowDirection2 = arrayOfInputs(44)
    parOperationLength = arrayOfInputs(45)
    parOperationHeight = arrayOfInputs(46)
    parOperationWidth = arrayOfInputs(47)
    parOperationWidth1 = arrayOfInputs(48)
    'Text Inputs
    parId1 = arrayOfInputs(49)
    parId2 = arrayOfInputs(50)
    parNPDUnitType = arrayOfInputs(51)
    parNPDUnitType1 = arrayOfInputs(52)
    parNPDUnitType2 = arrayOfInputs(53)
    
    iOutput = 0
    
    Dim MaxRadius As Double
    Dim MaxHeight As Double
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
'Insert your code for output 1(Insulation Body)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
'Maximum diameter amoung the three cylinders will be utilized for calculating the Insulation body width
'width .
    If (parInstrumentRadius >= parInstrumentRadius1) And (parInstrumentRadius >= parInstrumentRadius2) Then
        MaxRadius = parInstrumentRadius
    ElseIf (parInstrumentRadius2 >= parInstrumentRadius1) Then
        MaxRadius = parInstrumentRadius2
    Else
        MaxRadius = parInstrumentRadius1
    End If
'Maximum height amoung the four cylinders will be utilized for calculating the Insulation body heigth
    If (parInstrumentHeight3 > parInstrumentHeight5) And (parInstrumentHeight3 > parInstrumentHeight7) _
                                                                                    And (parInstrumentHeight3 > parInstrumentHeight9) Then
        MaxHeight = parInstrumentHeight3
    ElseIf (parInstrumentHeight5 > parInstrumentHeight7) And (parInstrumentHeight5 > parInstrumentHeight9) Then
        MaxHeight = parInstrumentHeight5
    ElseIf (parInstrumentHeight7 > parInstrumentHeight9) Then
        MaxHeight = parInstrumentHeight7
    Else
        MaxHeight = parInstrumentHeight9
    End If
    
    stPoint.Set parInstrumentLength + parInsulationThickness, _
                                parInstrumentHeight1 + MaxHeight + parInsulationThickness, _
                                -(parInstrumentWidth + MaxRadius * 2 + parInsulationThickness)
    enPoint.Set -parFace1toCenter, -parFace2toCenter, _
                    (parInstrumentWidth1 + parInstrumentRadius3 * 2 + parInsulationThickness)
    
    Set ObjInsBody = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    
    Set ObjInsBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
