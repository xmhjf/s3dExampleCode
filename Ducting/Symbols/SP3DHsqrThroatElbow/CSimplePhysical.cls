VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph
'   All Rights Reserved
'
'   CSimplePhysical.cls
'   Author:         NNC
'   Creation Date:  Wednesday, Feb 4 2004
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'       HVAC Square Throat Elbow
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Initialize()
'
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CP As New AutoMath.DPosition
    Dim NozzleLength As Double
    
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parThroat1 As Double
    Dim parThroat2 As Double
    Dim parPlaneOfTurn As Long
    Dim parCornerRadius As Double
    Dim parHVACShape As Integer
    
    Dim PortDirection As AutoMath.DVector
    Set PortDirection = New AutoMath.DVector
    Dim RadialDirection As AutoMath.DVector
    Set RadialDirection = New AutoMath.DVector

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parThroat1 = arrayOfInputs(4)
    parThroat2 = arrayOfInputs(5)
    parPlaneOfTurn = arrayOfInputs(6)
    
    iOutput = 0
    
    parHVACShape = 1
    parCornerRadius = 0

' Insert your code for output (Body)

' Place Nozzle 1
    PortDirection.Set -1, 0, 0
    If parPlaneOfTurn = 1 Then
        CP.Set -(parThroat1 + parWidth / 2), 0, 0
        RadialDirection.Set 0, 0, 1
        NozzleLength = parWidth + parThroat1
    Else
        CP.Set -(parThroat1 + parDepth / 2), 0, 0
        RadialDirection.Set 0, 1, 0
        NozzleLength = parDepth + parThroat1
    End If
    Dim oHvacNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle

    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort

    'NN: add HVAC parameters for Nozzle
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim iPortIndex As Integer

    iPortIndex = 1
    EndPrep = 11
    PortDepth = 0#
    CptOffset = 0#
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlangeWidth = 0#
    Thickness = 0.0001
   
    FlowDir = DistribFlow_BOTH
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            parHVACShape, EndPrep, _
                            Thickness, FlangeWidth, FlowDir, parWidth, _
                            parDepth, parCornerRadius, DimBaseOuter, PortStatus, _
                            "HvacNozzle1", PortDepth, CptOffset, False, _
                            m_OutputColl.ResourceManager)

    Set NozzleFactory = Nothing
    
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation CP
     
'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    iDistribPort.SetDirectionVector PortDirection

    iDistribPort.SetRadialOrient RadialDirection

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

' Place Nozzle 2
    PortDirection.Set 0, 1, 0
    NozzleLength = parThroat2
    If parPlaneOfTurn = 1 Then
        CP.Set 0, parThroat2 + parWidth / 2, 0
        RadialDirection.Set 0, 0, 1
    Else
        CP.Set 0, parThroat2 + parDepth / 2, 0
        RadialDirection.Set -1, 0, 0
    End If
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                    parHVACShape, EndPrep, _
                    Thickness, FlangeWidth, FlowDir, parWidth, _
                    parDepth, parCornerRadius, DimBaseOuter, PortStatus, _
                    "HvacNozzle2", PortDepth, CptOffset, False, _
                    m_OutputColl.ResourceManager)

     Set NozzleFactory = Nothing

     Set iDistribPort = oHvacNozzle
     iDistribPort.SetPortLocation CP

'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    iDistribPort.SetDirectionVector PortDirection

    iDistribPort.SetRadialOrient RadialDirection
'
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CP = Nothing
    Set PortDirection = Nothing
    Set RadialDirection = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

