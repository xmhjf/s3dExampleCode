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
'   ProgID:         SP3DHRectCross.HRectCross
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'    Definition of HVAC Rectangular-Reducing Cross
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI           As Double


Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHLength As Double
    Dim parBLength As Double
    Dim parInsulationThickness As Double
    Dim parPlaneOfBranch As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHLength = arrayOfInputs(6)
    parBLength = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    parPlaneOfBranch = arrayOfInputs(9)
    
    PI = 4 * Atn(1)

    iOutput = 0

'   If this symbol is placed on a round duct then Depth parameter is ZERO.
'   To guard the symbol depth is made equal to width.
    If parDepth = 0 Then parDepth = parWidth
    If parBDepth = 0 Then parBDepth = parBWidth

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

' Define default minimum Tee or Reducing Tee header and branch lengths
    dMinHL = parWidth + 0.05
    dMinBL = dMinHL / 2
    If CmpDblLessThan(parHLength, dMinHL) Then
        parHLength = dMinHL
    End If
    If CmpDblLessThan(parBLength, dMinBL) Then
        parBLength = dMinBL
    End If

'Set HVAC nozzle parameters
    iPortIndex = 1
    lEndPrep = 301
    dCornerRadius = 0#
    dNozzleLength = parHLength
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
                              lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parWidth, _
                              parDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                              "HvacNozzle1", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the conenct point of the nozzle
    oPos.Set -parHLength / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which Duct will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 0, 1
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
                            lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parWidth, _
                            parDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle2", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the connect point of the nozzle
    dNozzleLength = 0.0002

    oPos.Set parHLength / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 1, 0, 0

    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 0, 1
    oDistribPort.SetRadialOrient oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oDistribPort = Nothing
    
' Insert your code for output 3(HvacPort3 of Tee and Reducing Tee)
    iPortIndex = 3
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, _
                            lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parBWidth, _
                            parBDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle3", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the conenct point of the nozzle
    dNozzleLength = parBLength '- parDepth / 2 ''Branch Nozzle strats from the top of the Header surface

    oPos.Set 0, parBLength * Cos(parPlaneOfBranch), parBLength * Sin(parPlaneOfBranch)

    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector

    oDir.Set 0, Cos(parPlaneOfBranch), Sin(parPlaneOfBranch)
    oRadialDir.Set Sin(parPlaneOfBranch), 0, Cos(parPlaneOfBranch)

    oDistribPort.SetDirectionVector oDir
    oDistribPort.SetRadialOrient oRadialDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    
    ' Insert your code for output 3(HvacPort3 of Tee and Reducing Tee)
    iPortIndex = 4
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, _
                            lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parBWidth, _
                            parBDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle4", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

'Position of the nozzle should be the conenct point of the nozzle
    dNozzleLength = parBLength '- parDepth / 2 ''Branch Nozzle strats from the top of the Header surface

    oPos.Set 0, parBLength * Cos(parPlaneOfBranch + PI), parBLength * Sin(parPlaneOfBranch + PI)

    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.


    oDir.Set 0, Cos(parPlaneOfBranch + PI), Sin(parPlaneOfBranch + PI)
    oRadialDir.Set Sin(parPlaneOfBranch + PI), 0, Cos(parPlaneOfBranch + PI)

    oDistribPort.SetDirectionVector oDir
    oDistribPort.SetRadialOrient oRadialDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength
    

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oRadialDir = Nothing
        
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
