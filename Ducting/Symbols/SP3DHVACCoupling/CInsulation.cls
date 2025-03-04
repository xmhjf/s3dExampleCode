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
'   ProgID:         SP3DHVACCoupling.HVACCoupling
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'    Insulation aspect of HVAC HVAC Coupling
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'  14.Mar.2008     VRK      TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation:    'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Const EPSILON = 0.000001

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double

    Dim ObjBodyIns As Object

    Dim parWidth As Double
    Dim parBWidth As Double
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
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    ' a) Coupling, default-143
    ' b) Coupling, with takeoff-144

    Select Case lPartDataBasis

        Case Is <= 1, 143, 144
            dblInsulationDia = parWidth + parInsulationThickness * 2
        
            stPoint.Set -parLength / 2, 0, 0
            enPoint.Set parLength / 2, 0, 0
        
            Set ObjBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dblInsulationDia, False)
            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
            Set ObjBodyIns = Nothing
            
        Case Else
        '
    End Select
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
