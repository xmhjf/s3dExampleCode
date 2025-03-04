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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Monday, Jan 13 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
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
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHLength As Double
    Dim parBLength As Double
'   Angle is measured from positive X-axis in anti-clockwise direction
    Dim parAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHLength = arrayOfInputs(6)
    parBLength = arrayOfInputs(7)
    parAngle = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

    Dim minHL As Double
    Dim minBL As Double
    Dim halfHLength As Double
    
'   Define default minimum Tee or Reducing Tee header and branch lengths
    minHL = parWidth + 0.05
    minBL = minHL / 2
    If parHLength < minHL Then
        parHLength = minHL
    End If
    If parBLength < minBL Then
        parBLength = minBL
    End If
    halfHLength = parHLength / 2

' Insert your code for output 1(HvacPort1 of Tee and Reducing Tee)
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim nozzleIndex As Integer
    Dim ObjHvacNozzle1 As GSCADNozzleEntities.HvacNozzle
    Dim ObjHvacNozzle2 As GSCADNozzleEntities.HvacNozzle
    Dim ObjHvacNozzle3 As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim CornerRadius As Double
    Dim NozzleLength As Double
    Dim FlangeWidth As Double
    Dim Thickness As Double
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    
'   Set HVAC nozzle parameters
    iPortIndex = 1
    
    EndPrep = 301
    CornerRadius = 0#
    NozzleLength = parHLength

'   To construct nozzle as crosssection only, use FlangeWidth of 0, and
'   some non-zero value for flange thickness
'   Flange thickness assumed to be negigible thickness
    FlangeWidth = 0#
    Thickness = 0.0001
    PortDepth = 0#
    CptOffset = 0#
    FlowDir = DistribFlow_BOTH
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set ObjHvacNozzle1 = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    X = -halfHLength
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = ObjHvacNozzle1
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = ObjHvacNozzle1
    iNozzle.Length = NozzleLength
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacNozzle1
    Set ObjHvacNozzle1 = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
' Insert your code for output 2(HvacPort2 of Tee and Reducing Tee)
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
    Set ObjHvacNozzle2 = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
   
'   Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = 0.0002

    
    X = halfHLength
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = ObjHvacNozzle2
    iDistribPort.SetPortLocation pos

'   Direction specified here of the nozzle should be the direction in which pipe will be routed.
'   Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = ObjHvacNozzle2
    iNozzle.Length = NozzleLength
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacNozzle2
    Set ObjHvacNozzle2 = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
' Insert your code for output 3(HvacPort3 of Tee and Reducing Tee)
    iPortIndex = 3
    FlowDir = DistribFlow_BOTH
    Set ObjHvacNozzle3 = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                    Thickness, FlangeWidth, FlowDir, parBWidth, _
                                    parBDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                    "HNoz3", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
   
'   Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = parBLength
  
    X = parBLength * Cos(parAngle)
    Y = parBLength * Sin(parAngle)
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = ObjHvacNozzle3
    iDistribPort.SetPortLocation pos

'   Direction specified here of the nozzle should be the direction in which pipe will be routed.
'   Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set Cos(parAngle), Sin(parAngle), 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set -Sin(parAngle), Cos(parAngle), 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = ObjHvacNozzle3
    iNozzle.Length = NozzleLength

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacNozzle3
    Set ObjHvacNozzle3 = Nothing
    
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
