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
'   Copyright (c) 2003-2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          ACM
'   Creation Date:  Wednesday, Jan 15 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   29.Jul.2003     svsmylav            TR 44043: New input parameter 'PlaneOfBranch' is used
'                                       to orient Nozzle3(Route will use it for correct placement).
'                                       Checks are added to ensure non-zero Depth values.
'   7.May.2004      svsmylav            TR 58982: Modified Nozzle3's radial direction to take PlaneOfBranch into account.
'   15.DEC.2004     MS                  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'   24.Feb.2006     svsmylav            TR 93922: Modified radial direction of three ports such that
'                                       the Tee is oriented in X-Y plane for PlaneOfBranch equal to zero or PI/2.
'   08.SEP.2006     KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.Nov.2007     RUK                 DI-CP·56531 Ability for custom weight and cg calculation for parametric standard components.
'   02.Feb.2008     VRK                 TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages

Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHLength As Double
    Dim parBLength As Double
    Dim parInsulationThickness As Double
    Dim parPlaneOfBranch As Double
    Dim parLegLength1 As Double
    Dim parLegLength2 As Double
    Dim parLegLength3 As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parWidth = arrayOfInputs(2)
    'parDepth = arrayOfInputs(3)
    'parBWidth = arrayOfInputs(4)
    'parBDepth = arrayOfInputs(5)
    'parHLength = arrayOfInputs(6)
    'parBLength = arrayOfInputs(7)
    'parInsulationThickness = arrayOfInputs(8)
    'parPlaneOfBranch = arrayOfInputs(9)
    'parLegLength1 = arrayOfInputs(10)
    'parLegLength2 = arrayOfInputs(11)
    'parLegLength3 = arrayOfInputs(12)

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

    Dim PortStatus As DistribPortStatus
    Dim bDimBaseOuter As Boolean
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    Dim dMinHL As Double
    Dim dMinBL As Double
    Dim dNozzleLength As Double

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    ' a) Rectangular Tee, default-124
    ' b) Rectangular Tee, specified by leg lengths-125

    Select Case lPartDataBasis

    Case Is <= 1, 124 'Rectangular Tee, default
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parBWidth = arrayOfInputs(4)
        parBDepth = arrayOfInputs(5)
        parHLength = arrayOfInputs(6)
        parBLength = arrayOfInputs(7)
        parPlaneOfBranch = arrayOfInputs(9)
        'Define default minimum Tee or Reducing Tee header and branch lengths
        dMinHL = parWidth + 0.05
        dMinBL = dMinHL / 2
        If parHLength < dMinHL Then
            parHLength = dMinHL
        End If
        If parBLength < dMinBL Then
            parBLength = dMinBL
        End If

    Case 125 'Rectangular Tee, specified by leg lengths
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parBWidth = arrayOfInputs(4)
        parBDepth = arrayOfInputs(5)
        parPlaneOfBranch = arrayOfInputs(9)
        parLegLength1 = arrayOfInputs(10)
        parLegLength2 = arrayOfInputs(11)
        parLegLength3 = arrayOfInputs(12)
        parHLength = parLegLength1 + parBWidth + parLegLength2
        If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
            parBLength = parLegLength3 + parDepth / 2
        Else
            parBLength = parLegLength3 + parWidth / 2
        End If

    Case Else
        GoTo ErrorLabel:    'Invalid specification.

    End Select
     
    ' HVAC parameters for Nozzle
    Dim DimBaseOuter As Boolean
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    
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
    dNozzleLength = parHLength
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Rectangular, lEndPrep(1), _
                            dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                            parDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle1", dPortDepth(1), dCptOffSet(1), False, _
                            m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    oPos.Set -parHLength / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which Duct will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
        oDir.Set 0, 1, 0
    Else
        oDir.Set 0, 0, 1
    End If

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
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Rectangular, lEndPrep(2), _
                            dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                            parDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle2", dPortDepth(2), dCptOffSet(2), False, _
                            m_OutputColl.ResourceManager)
                            
    'Position of the nozzle should be the connect point of the nozzle
    dNozzleLength = 0

    oPos.Set parHLength / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

    If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
        oDir.Set 0, 1, 0
    Else
        oDir.Set 0, 0, 1
    End If
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
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Rectangular, lEndPrep(3), _
                            dThickness(3), dFlangeWidth(3), lFlowDir(3), parBWidth, _
                            parBDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                            "HvacNozzle3", dPortDepth(3), dCptOffSet(3), False, _
                            m_OutputColl.ResourceManager)

    If lPartDataBasis <= 1 Or lPartDataBasis = 124 Then
        'Position of the nozzle should be the conenct point of the nozzle
        dNozzleLength = parBLength    '- parDepth / 2 ''Branch Nozzle strats from the top of the Header surface
        oPos.Set 0, parBLength, 0
    End If
    If lPartDataBasis = 125 Then
        'Position of the nozzle should be the conenct point of the nozzle
        If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
            oPos.Set 0, parDepth / 2 + parLegLength3, 0
            dNozzleLength = parDepth / 2 + parLegLength3    '- parDepth / 2 ''Branch Nozzle strats from the top of the Header surface
        Else
            oPos.Set 0, parWidth / 2 + parLegLength3, 0
            dNozzleLength = parWidth / 2 + parLegLength3    '- parDepth / 2 ''Branch Nozzle strats from the top of the Header surface
        End If
    End If

    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector

    If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
        oDir.Set 0, 1, 0
        oRadialDir.Set 1, 0, 0
    Else
        oDir.Set 0, 1, 0
        oRadialDir.Set 0, 0, 1
    End If

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

    'New code for WCGOutputs DI-CP·56531
    Dim oGenBOFactory As IJGeneralBusinessObjectsFactory
    Dim oAttributes As IJDAttributes
    Dim oAttribCol As IJDAttributesCol
    Set oGenBOFactory = New GeneralBusinessObjectsFactory

    Set oAttributes = oGenBOFactory.CreateWCGPropHolder(m_OutputColl.ResourceManager)

    'In V8 we should use the IJWCGPropHolder interface to set these values
    Set oAttribCol = oAttributes.CollectionOfAttributes(CVar("IJWCGPropHolder"))

    Dim dTotVolume As Double
    Dim dHeaderVol As Double
    Dim dBranchVol As Double
    Dim dCGX As Double, dCGY As Double, dCGZ As Double

    'Calculate the Volume
    'Assumed that the thickness of duct is negligible
    dHeaderVol = 2 * (parWidth + parDepth) * parHLength - (parBWidth * parBDepth)
    If parPlaneOfBranch = 0 Then
        dBranchVol = 2 * (parBWidth + parBDepth) * (parBLength - parWidth / 2)
    Else
        dBranchVol = 2 * (parBWidth + parBDepth) * (parBLength - parDepth / 2)
    End If
    dTotVolume = dHeaderVol + dBranchVol

    dCGX = 0    'as this symbol is symmetric about X-axis this is zero
    If parPlaneOfBranch = 0 Then
        dCGY = (dBranchVol * (parWidth / 2 + (parBLength - parWidth / 2) / 2)) / dTotVolume
    Else
        dCGY = (dBranchVol * (parWidth / 2 + (parBLength - parDepth / 2) / 2)) / dTotVolume
    End If
    dCGZ = 0    'as this symbol is symmetric about Z-axis this is zero

    oAttribCol.Item("Volume").Value = dTotVolume    '(cubic meters)
    oAttribCol.Item("CGX").Value = dCGX
    oAttribCol.Item("CGY").Value = dCGY
    oAttribCol.Item("CGZ").Value = dCGZ

    m_OutputColl.AddOutput "StrippedWCG", oAttributes
    Set oAttribCol = Nothing
    Set oAttributes = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
