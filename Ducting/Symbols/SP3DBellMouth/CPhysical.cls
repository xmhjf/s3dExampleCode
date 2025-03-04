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

'   Author:         NN
'   Creation Date:  Friday 1, Aug 2003
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Hvac Bell Mouth.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
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
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBellMouth As Object

    Dim Width As Double
    Dim ThroatRadius As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    ThroatRadius = arrayOfInputs(3)

    iOutput = 0

 ' Insert your code for output 1(Bell Mouth)
    Dim objBell As IJDObject
    Dim PI As Double
    Dim BellStart As IJDPosition
    Dim BellEnd As IJDPosition
    Dim BellCenter As IJDPosition
    Dim Axis As IJDVector

    Set BellStart = New DPosition
    Set BellEnd = New DPosition
    Set BellCenter = New DPosition
    Set Axis = New DVector

    BellStart.Set CenterPos.X, CenterPos.Y + Width / 2, CenterPos.Z
    BellEnd.Set CenterPos.X + ThroatRadius, CenterPos.Y + Width / 2 + ThroatRadius, CenterPos.Z
    BellCenter.Set CenterPos.X, CenterPos.Y + Width / 2 + ThroatRadius, CenterPos.Z

    Set objBell = PlaceTrArcByCenter(BellStart, BellEnd, BellCenter)
    PI = 4 * Atn(1)
    Axis.Set 1, 0, 0
    Set ObjBellMouth = PlaceRevolution(m_OutputColl, objBell, Axis, CenterPos, PI * 2, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBellMouth
    Set ObjBellMouth = Nothing
    objBell.Remove
    Set objBell = Nothing
' Place Bell Mouth Port

'=====================================
'BUILD HVAC NOZZLE: Bell Mouth Port
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim Depth As Double
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim cptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double

    'Set HVAC nozzle parameters
    iPortIndex = 1
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
    PortDepth = 0#
    cptOffset = 0#
    Depth = Width
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness


    FlowDir = DistribFlow_IN

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "BellPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir

    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CenterPos = Nothing
    Set BellStart = Nothing
    Set BellEnd = Nothing
    Set BellCenter = Nothing
    Set Axis = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

