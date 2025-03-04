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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'   CPhysical.cls
'   Author:         kkk
'   Creation Date:  Thursday 23, March 2006
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for Physical aspect of
'       Hvac Bell Mouth.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   08.SEP.2006     KKC               DI-95670  Replace names with initials in all revision history sheets and symbols
'   20.SEP.2006     dkl/svsmylav      TR-104419: Removed 3 ObjBellPortion outputs as now we used skinned surface.
'   27.Nov.2006     svsmylav          TR-110632: Used 0.2 mm distance between port 1 and port 2 to avoid
'                                     zero length vector while using Common App Rotate command.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_VALUE = 0.0001

Private PI       As Double
Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
' Define the variables that are supplied in the catalog
    Dim Width As Double
    Dim Depth As Double
    Dim BWidth As Double
    Dim BDepth As Double
    Dim Length As Double
    Dim PlaneOfBranch As Double
    
' Define the variables that can be computed from the given dimensions
    Dim ThroatRadius As Double
    Dim HeaderWidth As Double
    Dim HeaderDepth As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    BWidth = arrayOfInputs(4)
    BDepth = arrayOfInputs(5)
    Length = arrayOfInputs(6)
    PlaneOfBranch = arrayOfInputs(7)
    
    ThroatRadius = Length - (25.4 / 1000) ' R = C - 1 Inch according to the catalog.
'   To simplify the symbol graphics, we did not include the tangents to the arc. Due to
'   this instead of actual value Header Width = Branch Width + 2 * Length in the drawing,
'   we need to use HeaderWidth = BWidth + (ThroatRadius * 2) ignoring the tangents
    HeaderWidth = BWidth + (ThroatRadius * 2)
    HeaderDepth = BDepth + (ThroatRadius * 2)
    
    iOutput = 0

 ' This componenent is being placed as a Branch Component)
 ' Though there is no Header graphics embedded inside the duct, only Nozzle-1 & 2 are being placed with
 ' minimal length just to make this component - Branch Type Component.
 
'=====================================
    'BUILD HVAC NOZZLE - 1: OutPut(1)
'=====================================
 Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
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
    
' To construct nozzle as crosssection only, use FlangeWidth of 0, and
' some non-zero value for flange thickness
' Flange thickness assumed to be negigible thickness

    FlowDir = DistribFlow_BOTH

' Depth and Width of crosssection will be the same as diameter of the nozzle which
' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -NEGLIGIBLE_VALUE
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
        dir.Set 0, 1, 0
    Else
        dir.Set 0, 0, 1
    End If

    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'=====================================
    'BUILD HVAC NOZZLE - 2:OutPut (2)
'=====================================
'Set HVAC nozzle parameters
    iPortIndex = 2
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
    PortDepth = 0#
    cptOffset = 0#
' To construct nozzle as crosssection only, use FlangeWidth of 0, and
' some non-zero value for flange thickness
' Flange thickness assumed to be negigible thickness

    FlowDir = DistribFlow_BOTH

' Depth and Width of crosssection will be the same as diameter of the nozzle which
' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
'Position of the nozzle should be the conenct point of the nozzle
    
    X = NEGLIGIBLE_VALUE
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir

    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
        dir.Set 0, 1, 0
    Else
         dir.Set 0, 0, 1
    End If
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

'=========================================================
'BUILD HVAC NOZZLE 3: For Branch -Reducing Bell Mouth Port
'=========================================================
'Set HVAC nozzle parameters
    iPortIndex = 3
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = (25.4 / 1000)
    PortDepth = 0#
    cptOffset = 0#
    
    FlowDir = DistribFlow_OUT
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth, _
                                            BDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
'Position of the nozzle should be the conenct point of the nozzle
    
    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
        X = 0#: Y = (Depth / 2) + ThroatRadius + NozzleLength: Z = 0#
    Else
        X = 0#: Y = (Width / 2) + ThroatRadius + NozzleLength: Z = 0#
    End If
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
    dir.Set 0, 1, 0
    iDistribPort.SetDirectionVector dir
    
    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
         dir.Set 1, 0, 0
    Else
          dir.Set 0, 0, 1
    End If
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

'=======================================
'BUILD Bell Mouth Portion
'=======================================
  
    Dim objCurve1 As IngrGeom3D.Arc3d

    Dim BellStart As IJDPosition
    Dim BellEnd As IJDPosition
    Dim BellCenter As IJDPosition
    
    Set BellStart = New DPosition
    Set BellEnd = New DPosition
    Set BellCenter = New DPosition

' Forming Curve - 1
    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
        BellStart.Set -HeaderDepth / 2, Depth / 2, 0
        BellEnd.Set -BDepth / 2, (Depth / 2) + ThroatRadius, 0
        BellCenter.Set -HeaderDepth / 2, (Depth / 2) + ThroatRadius, 0
    Else
        BellStart.Set -HeaderWidth / 2, Width / 2, 0
        BellEnd.Set -BWidth / 2, (Width / 2) + ThroatRadius, 0
        BellCenter.Set -HeaderWidth / 2, (Width / 2) + ThroatRadius, 0
    End If
    Set objCurve1 = PlaceTrArcByCenter(BellStart, BellEnd, BellCenter)
    
'   Prepare Line Strings
    Dim oTopLineString As IngrGeom3D.LineString3d
    Dim oBotLineString As IngrGeom3D.LineString3d
    Dim dLineStrPoints(0 To 17)  As Double
    
'   Prepare Bottom Line String
    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
        dLineStrPoints(0) = -HeaderDepth / 2
        dLineStrPoints(1) = Depth / 2
        dLineStrPoints(2) = 0
    
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = HeaderWidth / 2

        dLineStrPoints(6) = HeaderDepth / 2
        dLineStrPoints(7) = dLineStrPoints(1)
        dLineStrPoints(8) = dLineStrPoints(5)

        dLineStrPoints(9) = dLineStrPoints(6)
        dLineStrPoints(10) = dLineStrPoints(1)
        dLineStrPoints(11) = -HeaderWidth / 2

        dLineStrPoints(12) = -HeaderDepth / 2
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(11)
        
        dLineStrPoints(15) = dLineStrPoints(0)
        dLineStrPoints(16) = dLineStrPoints(1)
        dLineStrPoints(17) = dLineStrPoints(2)
    Else
        dLineStrPoints(0) = -HeaderWidth / 2
        dLineStrPoints(1) = Width / 2
        dLineStrPoints(2) = 0
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = HeaderDepth / 2
        
        dLineStrPoints(6) = HeaderWidth / 2
        dLineStrPoints(7) = dLineStrPoints(1)
        dLineStrPoints(8) = dLineStrPoints(5)
        
        dLineStrPoints(9) = dLineStrPoints(6)
        dLineStrPoints(10) = dLineStrPoints(1)
        dLineStrPoints(11) = -HeaderDepth / 2
        
        dLineStrPoints(12) = -HeaderWidth / 2
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(11)
        
        dLineStrPoints(15) = dLineStrPoints(0)
        dLineStrPoints(16) = dLineStrPoints(1)
        dLineStrPoints(17) = dLineStrPoints(2)
    End If
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oBotLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dLineStrPoints)
    
'   Prepare Top Line String
    If CmpDblEqual(PlaneOfBranch, PI / 2) Then
        dLineStrPoints(0) = -BDepth / 2
        dLineStrPoints(1) = Depth / 2 + ThroatRadius
        dLineStrPoints(2) = 0
    
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = BWidth / 2

        dLineStrPoints(6) = BDepth / 2
        dLineStrPoints(7) = dLineStrPoints(1)
        dLineStrPoints(8) = dLineStrPoints(5)

        dLineStrPoints(9) = dLineStrPoints(6)
        dLineStrPoints(10) = dLineStrPoints(1)
        dLineStrPoints(11) = -BWidth / 2

        dLineStrPoints(12) = -BDepth / 2
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(11)
'
        dLineStrPoints(15) = dLineStrPoints(0)
        dLineStrPoints(16) = dLineStrPoints(1)
        dLineStrPoints(17) = dLineStrPoints(2)
    Else
        dLineStrPoints(0) = -BWidth / 2
        dLineStrPoints(1) = (Width / 2) + ThroatRadius
        dLineStrPoints(2) = 0
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = BDepth / 2
        
        dLineStrPoints(6) = BWidth / 2
        dLineStrPoints(7) = dLineStrPoints(1)
        dLineStrPoints(8) = dLineStrPoints(5)
        
        dLineStrPoints(9) = dLineStrPoints(6)
        dLineStrPoints(10) = dLineStrPoints(1)
        dLineStrPoints(11) = -BDepth / 2
        
        dLineStrPoints(12) = -BWidth / 2
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(11)
        
        dLineStrPoints(15) = dLineStrPoints(0)
        dLineStrPoints(16) = dLineStrPoints(1)
        dLineStrPoints(17) = dLineStrPoints(2)
    End If
    Set oTopLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dLineStrPoints)

    Dim trObj As IJElements
    Dim csObj As IJElements
    Set trObj = New JObjectCollection
    Set csObj = New JObjectCollection
    
    csObj.Add objCurve1
    trObj.Add oBotLineString
    trObj.Add oTopLineString
    
    Dim Surfset   As IngrGeom3D.IJElements
    Dim WCaps As Long
    Set Surfset = geomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
        trObj, csObj, WCaps)
    Set geomFactory = Nothing
    Set oBotLineString = Nothing
    Set oTopLineString = Nothing
    Dim ObjBellMouthSurface As Object
    Dim iCount As Long
    For Each ObjBellMouthSurface In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBellMouthSurface
    Next ObjBellMouthSurface
    Set Surfset = Nothing
    Set ObjBellMouthSurface = Nothing
    Set objCurve1 = Nothing
    Set BellStart = Nothing
    Set BellEnd = Nothing
    Set BellCenter = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
