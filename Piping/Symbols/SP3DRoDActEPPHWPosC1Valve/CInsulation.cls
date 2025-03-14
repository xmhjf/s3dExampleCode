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
'   Author:          BG
'   Creation Date:  Wednesday, Oct 9 2002
'   Description:
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVJU).
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVJU).
'    this Symbol is same as (PDMS Category Name=IVJS)
'   Physical aspect consist 2 cones, 2 Cylinders of stem Dia, 2 Cylinder , 1 box, 1 torus and a cone with elliptical dome.
'   2 Nozzles. Insulation Aspect consist of a simple Cylinder for the valve Body.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub
Private Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue
End Function
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjValveBodyIns  As Object
    
    Dim parFace1toCenter As Double         'D2
    Dim parFacetoFace     As Double     'D3
    Dim parCylHeight      As Double     'D4
    Dim parActuatorWidth     As Double      'D5
    Dim parActuatorDiameter     As Double       'D6
    Dim parActuatorHeight     As Double          'D7
    Dim parHandWheelLength    As Double         'D8
    Dim parHandWheelDiameter      As Double         'D9
    Dim parActuatorHeight1     As Double            'D10
    Dim parActuatorHeight2     As Double            'D11
    Dim parActuatorLength     As Double             'D12
    Dim parActuatorLength1    As Double             'D14
    Dim parActuatorHeight3     As Double            'D16
    Dim parActuatorHeight4     As Double            'D19
    Dim parHandWheelOffset   As Double          'D20
    Dim parActuatorWidth1       As Double           'D25
    Dim parActuatorWidth2     As Double             'D26
    Dim parActuatorHeight5     As Double            'D27
    Dim parActuatorLength2  As Double           'D28
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)         'D2
    parFacetoFace = arrayOfInputs(3)         'D3
    parCylHeight = arrayOfInputs(4)         'D4
    parActuatorWidth = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)    'D6
    parActuatorHeight = arrayOfInputs(7)    'D7
    parHandWheelLength = arrayOfInputs(8)    'D8
    parHandWheelDiameter = arrayOfInputs(9)     'D9
    parActuatorHeight1 = arrayOfInputs(10)    'D10
    parActuatorHeight2 = arrayOfInputs(11)    'D11
    parActuatorLength = arrayOfInputs(12)    'D12
    parActuatorLength1 = arrayOfInputs(13)    'D14
    parActuatorHeight3 = arrayOfInputs(14)    'D16
    parActuatorHeight4 = arrayOfInputs(15)    'D19
    parHandWheelOffset = arrayOfInputs(16)    'D20
    parActuatorWidth1 = arrayOfInputs(17)    'D25
    parActuatorWidth2 = arrayOfInputs(18)    'D26
    parActuatorHeight5 = arrayOfInputs(19)    'D27
    parActuatorLength2 = arrayOfInputs(20)    'D28
    parInsulationThickness = arrayOfInputs(21)
    
    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 8(Valve Body Insulation)
    Dim CylEdgeToCen As Double
    Dim CylDia As Double
    CylDia = parFace1toCenter
    CylEdgeToCen = Sqr(parCylHeight ^ 2 + (CylDia / 2) ^ 2)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxRad refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam1 / 2, pipeDiam2 / 2, flangeDiam1 / 2, flangeDiam2 / 2, CylEdgeToCen)
    InsulationDia = 2 * maxRad + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

