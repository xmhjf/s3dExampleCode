VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday 1, Aug 2003
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Symbol Hvac Surface Mounted component with one Round nozzle(Code is taken from
'       NN's SP3DHvacSMRB and modified).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   19, Oct. 2003   NN   Remove BranchOffset input, as Route controls the component offset.
'                                   Define     parBranchDepth = parBranchWidth for Round branch duct.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim parBranchWidth As Double
    Dim parBranchDepth As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBranchWidth = arrayOfInputs(2)

    iOutput = 0

'============================================================
'BUILD HVAC NOZZLE
'============================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim CornerRadius As Double
    Dim NozzleLength As Double
    Dim FlangeWidth As Double
    Dim Thickness As Double

    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer

    'Set HVAC nozzle parameters
    iPortIndex = 1

    EndPrep = 301
    CornerRadius = 0#
    NozzleLength = 0.0002

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlangeWidth = 0#
    Thickness = 0.0001
    PortDepth = 0#
    CptOffset = 0#
    FlowDir = DistribFlow_IN
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    parBranchDepth = parBranchWidth
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, parBranchWidth, _
                                            parBranchDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HvacSMPort", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)

    'Position of the nozzle should be the connect point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 0, 0, 1
    iDistribPort.SetDirectionVector dir

    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    Set iDistribPort = Nothing
    Set iNozzle = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

