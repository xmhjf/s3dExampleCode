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
' Symbol places Rivets for Flat Oval ducts.
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
    Dim RivetDiameter As Double
    Dim centerPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    
'   If this symbol is placed on a round duct, then parameter parDepth is ZERO.
'   'CreateHvacNozzle' method requires non-zero width and depth values.
    If parDepth = 0 Then parDepth = parWidth
    
    Dim dNozzleWidth As Double, dNozzleDepth As Double
    dNozzleWidth = parWidth ' use these for nozzle creation
    dNozzleDepth = parDepth
    
    iOutput = 0
    
'   Rivet diameter is taken as 8mm and this can be adjusted as required
    RivetDiameter = 0.008
    
'   If parWidth and parDepth are Not equal use logic for flat-oval duct rivets
    If parWidth <> parDepth Then
    '   Flat oval shape duct needs this change in parWidth
        parWidth = parWidth - parDepth
    
    ' Insert your code for output 1(Rivets)
        Dim I As Integer
        Dim HD As Double    'Half depth
    
        HD = parDepth / 2
                                     
    '   Flat Oval duct shape need determination of number of rivets per width
    '   Assumptions:
    '       Gap between any two adjacent rivets = 8 * rivet diameter
    '       Minimum number of rivets along width = 4
    '       Maximum number of rivets along width = 8
        Dim RivetSpacingAlongWidth  As Double
        
        NumRivetsAlongWidth = parWidth / (RivetDiameter * 8)
    
        If NumRivetsAlongWidth < 4 Then
            NumRivetsAlongWidth = 4
        ElseIf NumRivetsAlongWidth > 8 Then
            NumRivetsAlongWidth = 8
        End If
    
        RivetSpacingAlongWidth = parWidth / NumRivetsAlongWidth
    
        Dim FOCurveRadius As Double 'Radius of curved surfaces
        FOCurveRadius = parDepth / 2
        
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
        
    '   For the curved surfaces at front and rear three rivets per surface at 60 degrees are considered
    '   I) Front surface
    
    '   Middle Rivet
        centerPoint.Set 0, 0, parWidth / 2 + FOCurveRadius
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
    '   Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
        
    '   Upper Rivet
        centerPoint.Set 0, FOCurveRadius * Sin(PI / 3), parWidth / 2 + FOCurveRadius * Cos(PI / 3)
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
    '   Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
        
    '   Lower Rivet
        centerPoint.Set 0, -FOCurveRadius * Sin(PI / 3), parWidth / 2 + FOCurveRadius * Cos(PI / 3)
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
    '   Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
        
    '   II) Rear surface
        centerPoint.Set 0, 0, -(parWidth / 2 + FOCurveRadius)
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
    '   Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
        
    '   Upper Rivet
        centerPoint.Set 0, FOCurveRadius * Sin(PI / 3), -(parWidth / 2 + FOCurveRadius * Cos(PI / 3))
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
    '   Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
        
    '   Lower Rivet
        centerPoint.Set 0, -FOCurveRadius * Sin(PI / 3), -(parWidth / 2 + FOCurveRadius * Cos(PI / 3))
        Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
    
    '   Set the output
        m_OutputColl.AddOutput "Rivets_", ObjRivet
        Set ObjRivet = Nothing
    Else
'       parWidth and parDepth are Equal, So use logic similar to round duct rivets symbol.
        For I = 1 To 6
            centerPoint.Set 0, (parWidth / 2) * Cos((I - 1) * (PI / 3)), _
                            (parWidth / 2) * Sin((I - 1) * (PI / 3))
            Set ObjRivet = PlaceSphere(m_OutputColl, centerPoint, RivetDiameter / 2)
            
    '       Set the output
            m_OutputColl.AddOutput CStr("Rivet" & I), ObjRivet
            Set ObjRivet = Nothing
        Next I
    End If
    
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

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.FlatOval, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, dNozzleWidth, _
                                            dNozzleDepth, CornerRadius, DimBaseOuter, PortStatus, _
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
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(2, "SymbDefn", GSCADNozzleEntities.FlatOval, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, dNozzleWidth, _
                                            dNozzleDepth, CornerRadius, DimBaseOuter, PortStatus, _
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
    

