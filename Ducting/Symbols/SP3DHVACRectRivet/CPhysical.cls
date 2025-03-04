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
'   Creation Date:  Wednesday, Dec 11 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Symbol places Rivets for Rectangular ducts.
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
Private PI       As Double
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjRivet As Object
    
    Dim parWidth As Double
    Dim parDepth As Double
    
    Dim NumRivetsAlongWidth As Integer
    Dim NumRivetsAlongDepth As Integer
    Dim RivetDiameter As Double
    Dim centerPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    
    iOutput = 0
    
'   Rivet diameter is taken as 8mm and this can be adjusted as required
    RivetDiameter = 0.008

' Insert your code for output 1(Rivets)
    Dim I As Integer
    Dim HW As Double    'Half width
    Dim HD As Double    'Half depth
    
    HW = parWidth / 2
    HD = parDepth / 2
                                            
'   Rectangle duct shape need determination of number of rivets per width and depth
'   Assumptions:
'       Gap between any two adjacent rivets = 8 * rivet diameter
'       Minimum number of rivets along width = 4
'       Minimum number of rivets along depth = 3
'       Maximum number of rivets on any side of the duct = 8
    Dim RivetSpacingAlongWidth  As Double
    Dim RivetSpacingAlongDepth  As Double
    
    NumRivetsAlongWidth = parWidth / (RivetDiameter * 8)
    NumRivetsAlongDepth = parDepth / (RivetDiameter * 8)
    
    If NumRivetsAlongWidth < 4 Then
        NumRivetsAlongWidth = 4
    ElseIf NumRivetsAlongWidth > 8 Then
        NumRivetsAlongWidth = 8
    End If
    
    If NumRivetsAlongDepth < 3 Then
        NumRivetsAlongDepth = 3
    ElseIf NumRivetsAlongDepth > 8 Then
        NumRivetsAlongDepth = 8
    End If
    
    RivetSpacingAlongWidth = parWidth / NumRivetsAlongWidth
    RivetSpacingAlongDepth = parDepth / NumRivetsAlongDepth

'   Prepare for Rivets along width - Top surface
    For I = 1 To NumRivetsAlongWidth
        centerPoint.Set 0, HD, (0.5 * (NumRivetsAlongWidth + 1) - I) * RivetSpacingAlongWidth
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
'       Set the output
        m_OutputColl.AddOutput CStr("Rivet" & I), ObjRivet
        Set ObjRivet = Nothing
    Next I

'   Prepare for Rivets along width - Bottom surface
    For I = 1 To NumRivetsAlongWidth
        centerPoint.Set 0, -HD, (0.5 * (NumRivetsAlongWidth + 1) - I) * RivetSpacingAlongWidth
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
'       Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
    Next I

'   Prepare for Rivets along depth - Front surface
    For I = 1 To NumRivetsAlongDepth
        centerPoint.Set 0, (0.5 * (NumRivetsAlongDepth + 1) - I) * RivetSpacingAlongDepth, HW
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
                   
'       Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
    Next I

'   Prepare for Rivets along depth - Rear surface
    For I = 1 To NumRivetsAlongDepth
        centerPoint.Set 0, (0.5 * (NumRivetsAlongDepth + 1) - I) * RivetSpacingAlongDepth, -HW
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
     
'       Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
    Next I

' Place Nozzle 1
    Dim FlangeWidth As Double
    Dim NozzleFlangeThickness As Double
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim RadialDir As New AutoMath.DVector
    Dim CornerRadius As Double
        
'   Set HVAC nozzle parameters
    PortDepth = 0#

'   FlangeWidth and Flange thickness assumed to be negigible thickness, Nozzle length to be
'   greater than flange thickness in general.
    FlangeWidth = NEGLIGIBLE_THICKNESS
    NozzleFlangeThickness = NEGLIGIBLE_THICKNESS
    
    FlowDir = DistribFlow_UNDEFINED
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    EndPrep = 301
    CptOffset = 0
    CornerRadius = 0

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -NEGLIGIBLE_THICKNESS
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

'   Direction specified here of the nozzle should be the direction in which pipe will be routed.
'   Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    RadialDir.Set 0, 1, 0
    iDistribPort.SetRadialOrient RadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 2 * NEGLIGIBLE_THICKNESS

    m_OutputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

' Place Nozzle 2
    EndPrep = 301
    FlowDir = DistribFlow_UNDEFINED
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(2, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac2", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
   
'   Position of the nozzle should be the conenct point of the nozzle
    Set iDistribPort = oHvacNozzle
    X = NEGLIGIBLE_THICKNESS
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    iDistribPort.SetPortLocation pos

    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    iDistribPort.SetRadialOrient RadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 2 * NEGLIGIBLE_THICKNESS

    m_OutputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    
    Set iDistribPort = Nothing
    Set centerPoint = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    Set RadialDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


