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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DHRec2RecAdapter.HRec2RecAdapter
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private RAD          As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim Width As Double
    Dim Depth As Double
    Dim Width2 As Double
    Dim Depth2 As Double
    Dim Length As Double
    Dim InsulationThickness As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Width2 = arrayOfInputs(4)
    Depth2 = arrayOfInputs(5)
    Length = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    
    
    PI = 4 * Atn(1)

    iOutput = 0

'   If this symbol is placed on a round duct then Depth parameter is ZERO.
'   To guard the symbol depth is made equal to width.
    If Depth = 0 Then Depth = Width
    If Depth2 = 0 Then Depth2 = Width2

' Insert your code for output 1(HvacPort1 of Tee and Reducing Tee)

    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim lEndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim bDimBaseOuter As Boolean
    Dim dPortDepth As Double
    Dim dCptOffset As Double
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    Dim dMinHL As Double
    Dim dMinBL As Double
    Dim dNozzleLength As Double
    Dim dFlangeWidth As Double
    Dim dFlangeThickness As Double

'Set HVAC nozzle parameters
    iPortIndex = 1
    lEndPrep = 301
    dCornerRadius = 0#
    dNozzleLength = Length / 2
'NozzleLength Has to be GREATER than NozzleFlangeThickness
'To construct nozzle as crosssection only, use FlangeWidth of 0, and
'some non-zero value for flange thickness
'Flange thickness assumed to be negigible thickness
    dFlangeWidth = 0#
    dFlangeThickness = 0.0001
    dPortDepth = 0#
    dCptOffset = 0#
    FlowDir = DistribFlow_BOTH
    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, _
                              lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, Width, _
                              Depth, dCornerRadius, bDimBaseOuter, PortStatus, _
                              "HvacNozzle1", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the conenct point of the nozzle
    oPos.Set -Length / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which Duct will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oDistribPort = Nothing

'Insert your code for output 2(HvacPort2 of Tee and Reducing Tee)
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, _
                            lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, Width2, _
                            Depth2, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle2", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the connect point of the nozzle
    dNozzleLength = Length / 2

    oPos.Set Length / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 1, 0, 0

    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oDistribPort = Nothing
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
        
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


