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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DHVACEndCap.HVACEndCap
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'   Physical aspect of HVAC End Cap
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  24.Jan.2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const INCH = 0.0254
Private Const MODULE = "Physical:"    'Used for error messages
Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim parWidth As Double
    Dim parLength As Double
    Dim parInsulationThickness As Double
    Dim ObjCapEnd As Object


    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parLength = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

    ' Insert your code for output 1(HvacPort1 )

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
    Dim dNozzleLength As Double
    Dim dFlangeWidth As Double
    Dim dFlangeThickness As Double
    Dim dCapEndProjLen As Double
    Dim dCapEndProjDia As Double

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    Select Case lPartDataBasis
'   a) EndCap, default-130
'   b) EndCap, with takeoff-131

    Case Is <= 1, 130 'EndCap, default
        ' Insert your code for output 1(Cap End Body)
        dCapEndProjLen = 0.002
        dCapEndProjDia = parWidth * 1.05
        stPoint.Set parLength - dCapEndProjLen, 0, 0
        enPoint.Set parLength, 0, 0
        Set ObjCapEnd = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCapEndProjDia, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCapEnd
        Set ObjCapEnd = Nothing

    Case 131  'EndCap, with takeoff
        ' Insert your code for output 1(Cap End Body)

        dCapEndProjDia = parWidth * 1.05
        stPoint.Set 0, 0, 0
        enPoint.Set parLength, 0, 0
        Set ObjCapEnd = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCapEndProjDia, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCapEnd
        Set ObjCapEnd = Nothing

        Dim dTakeOff As Double
        Dim objTakeOff As Object
        dTakeOff = 2 * INCH   '2in
        stPoint.Set 0, 0, 0
        enPoint.Set -dTakeOff, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth * 1.01, True)
        ' Set the output
        m_OutputColl.AddOutput "ObjTakeOff", objTakeOff
        Set objTakeOff = Nothing

    Case Else
        GoTo ErrorLabel:    'Invalid specification.

    End Select


    'Set HVAC nozzle parameters
    iPortIndex = 1
    lEndPrep = 301
    dCornerRadius = 0#

    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    'To construct nozzle as crosssection only, use FlangeWidth of 0, and
    'some non-zero value for flange thickness
    'Flange thickness assumed to be negigible thickness
    dFlangeWidth = 0#
    dFlangeThickness = 0
    dPortDepth = 0#
    dCptOffset = 0#
    FlowDir = DistribFlow_BOTH
    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, _
                                                      lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parWidth, _
                                                      parWidth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                                      "HvacNozzle1", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

    'Check for partdata basi scases
    If lPartDataBasis = 131 Then
        dNozzleLength = 0
    Else
        dNozzleLength = parLength - dCapEndProjLen
    End If

    'Position of the nozzle should be the conenct point of the nozzle
    oPos.Set 0, 0, 0
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
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
