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
'   ProgID:         SP3DHVACCoupling.HVACCoupling
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'    Physical aspect of HVAC Coupling
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'  14.Mar.2008     VRK      TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
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
    Dim ObjTaperBody1 As Object
    Dim ObjTaperBody2 As Object

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parLength = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim dTaperHalfLength As Double

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    Select Case lPartDataBasis
'   a) Coupling, default-143
'   b) Coupling, with takeoff-144

    Case Is <= 1, 143 'Coupling, default

        ' Insert your code for output 1(Taper Body 1)
        dTaperHalfLength = 0.015
        stPoint.Set -dTaperHalfLength / 2, 0, 0
        enPoint.Set 0, 0, 0
        Set ObjTaperBody1 = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, (parWidth / 2) * 0.9, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperBody1
        Set ObjTaperBody1 = Nothing

        ' Insert your code for output 2(Taper Body 2)
        stPoint.Set dTaperHalfLength / 2, 0, 0
        enPoint.Set 0, 0, 0
        Set ObjTaperBody2 = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, (parWidth / 2) * 0.9, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperBody2
        Set ObjTaperBody2 = Nothing

    Case 144 'Coupling, with takeoff

        ' Insert your code for output 1(Taper Body 1)
        dTaperHalfLength = 0
        stPoint.Set -parLength / 2, 0, 0
        enPoint.Set 0, 0, 0
        Set ObjTaperBody1 = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, (parWidth / 2) * 1.05, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperBody1
        Set ObjTaperBody1 = Nothing

        ' Insert your code for output 2(Taper Body 2)
        stPoint.Set parLength / 2, 0, 0
        enPoint.Set 0, 0, 0
        Set ObjTaperBody2 = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, 1.05 * parWidth / 2, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperBody2
        Set ObjTaperBody2 = Nothing

        'Input for Takeoff
        Dim objTakeOff As Object
        Dim dTakeOff As Double
        dTakeOff = 2 * INCH   '2in
        stPoint.Set -parLength / 2, 0, 0
        enPoint.Set (-parLength / 2) - dTakeOff, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth * 1.01, True)
        ' Set the output
        m_OutputColl.AddOutput "ObjTakeOff1", objTakeOff
        Set objTakeOff = Nothing

        stPoint.Set parLength / 2, 0, 0
        enPoint.Set (parLength / 2) + dTakeOff, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth * 1.01, True)
        ' Set the output
        m_OutputColl.AddOutput "ObjTakeOff2", objTakeOff
        Set objTakeOff = Nothing

    Case Else
        GoTo ErrorLabel:    'Invalid specification.

    End Select

    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    Dim dNozzleLength As Double

    ' HVAC parameters for Nozzle
    Dim PortStatus As DistribPortStatus
    Dim bDimBaseOuter As Boolean
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection

    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double

    dCornerRadius = 0#
    'Set HVAC nozzle parameters from Catalog
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
    Next iPortIndex

    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing

    iPortIndex = 1
    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    dNozzleLength = (parLength / 2) - (dTaperHalfLength / 2)

    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                                      GSCADNozzleEntities.Round, lEndPrep(1), _
                                                      dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                                      parWidth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                                      "HvacNozzle1", dPortDepth(1), dCptOffSet(1), False, _
                                                      m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    oPos.Set -parLength / 2, 0, 0
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
    

    'Insert your code for output 3(HVAC Nozzle - 2)
    iPortIndex = 2
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                                      GSCADNozzleEntities.Round, lEndPrep(2), _
                                                      dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                                                      parWidth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                                      "HvacNozzle2", dPortDepth(2), dCptOffSet(2), False, _
                                                      m_OutputColl.ResourceManager)

    'Position of the nozzle should be the connect point of the nozzle
    oPos.Set parLength / 2, 0, 0
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
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
