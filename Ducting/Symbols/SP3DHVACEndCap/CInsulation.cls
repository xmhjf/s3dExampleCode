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
'   ProgID:         SP3DHVACEndCap.HVACEndCap
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'   Insulation aspect of HVAC Round Cap
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  24.Jan.2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const INCH = 0.0254
Private Const MODULE = Insulation:    'Used for error messages
Private Sub Class_Initialize()
''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double

    Dim ObjBodyIns As Object

    Dim parWidth As Double
    Dim parLength As Double
    Dim parInsulationThickness As Double
    Dim dblInsulationDia As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parLength = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    dblInsulationDia = parWidth + parInsulationThickness * 2

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Select Case lPartDataBasis

    Case Is <= 1, 130 'EndCap, default
        ' Insert your code for output 1(Cap End Body)
        stPoint.Set 0, 0, 0
        enPoint.Set parLength, 0, 0

        Set ObjBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dblInsulationDia, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
        Set ObjBodyIns = Nothing

    Case 131 'EndCap, with takeoff
        ' Insert your code for output 1(Cap End Body)
        stPoint.Set 0, 0, 0
        enPoint.Set parLength, 0, 0

        Set ObjBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dblInsulationDia, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
        Set ObjBodyIns = Nothing

        Dim dTakeOff As Double
        Dim objTakeOff As Object
        dTakeOff = 2 * INCH   '2in
        stPoint.Set 0, 0, 0
        enPoint.Set -dTakeOff, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                       1.01 * parWidth + 2 * parInsulationThickness, True)
        ' Set the output
        m_OutputColl.AddOutput "ObjInsTakeOff", objTakeOff
        Set objTakeOff = Nothing

    Case Else
        GoTo ErrorLabel:    'Invalid specification.

    End Select

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
