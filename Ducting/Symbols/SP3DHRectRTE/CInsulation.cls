VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          ACM
'   Creation Date:  Wednesday, Jan 15 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     -----        ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   29.Jul.2003     svsmylav            TR 44043: New input parameter 'PlaneOfBranch' is used
'                                       to orient Nozzle3(Route will use it for correct placement).
'                                       Checks are added to ensure non-zero Depth values.
'   7.May.2004      svsmylav            TR 58982: Modified insulation code to take care of
'                                       change in Nozzle3's radial direction (Modified Nozzle3's
'                                       radial direction to take PlaneOfBranch into account).
'   24.Feb.2006     svsmylav            TR 93922: Modified the symbol such that the Tee is oriented
'                                       in X-Y plane for PlaneOfBranch equal to zero or PI/2.
'   08.SEP.2006     KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.Nov.2007     RUK                 DI-CP·56531 Ability for custom weight and cg calculation for parametric standard components.
'   02.Feb.2008     VRK                 TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation:    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Const EPSILON = 0.000001

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double

    Dim ObjHvacHeaderIns As Object
    Dim ObjHvacBranchIns As Object

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
    parInsulationThickness = arrayOfInputs(8)
    'parPlaneOfBranch = arrayOfInputs(9)
    'parLegLength1 = arrayOfInputs(10)
    'parLegLength2 = arrayOfInputs(11)
    'parLegLength3 = arrayOfInputs(12)

    iOutput = 0

    '   If this symbol is placed on a round duct then Depth parameter is ZERO.
    '   To guard the symbol depth is made equal to width.
    If parDepth = 0 Then parDepth = parWidth
    If parBDepth = 0 Then parBDepth = parBWidth

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
        Dim dMinHL As Double
        Dim dMinBL As Double

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

    ' Insert your code for output 4(Hvac Header Insulation)
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = -parHLength / 2
    dY = parWidth / 2 + parInsulationThickness
    dZ = parDepth / 2 + parInsulationThickness

    If CmpDblEqual(parPlaneOfBranch, PI / 2) Then
        dY = parDepth / 2 + parInsulationThickness
        dZ = parWidth / 2 + parInsulationThickness
    Else
        dY = parWidth / 2 + parInsulationThickness
        dZ = parDepth / 2 + parInsulationThickness
    End If

    Dim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = dX
    dlinepts(7) = -dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = dX
    dlinepts(10) = -dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Dim oAxisVec As New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    Set ObjHvacHeaderIns = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, parHLength, True)

    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacHeaderIns
    Set ObjHvacHeaderIns = Nothing

    ' Insert your code for output 5(Hvac Branch Insulation)
    Dim dProjLength As Double
    If Abs(parPlaneOfBranch) < EPSILON Or Abs((Abs(parPlaneOfBranch) - PI)) < EPSILON Then
        dX = (parBWidth / 2 + parInsulationThickness)
        dY = parBLength
        dZ = (parBDepth / 2 + parInsulationThickness)
        dProjLength = parBLength - parWidth / 2 - parInsulationThickness
    Else
        dX = (parBDepth / 2 + parInsulationThickness)
        dY = parBLength
        dZ = (parBWidth / 2 + parInsulationThickness)
        dProjLength = parBLength - parDepth / 2 - parInsulationThickness
    End If

    'Ensure minimum projection length for third nozzle insulation
    If dProjLength < EPSILON Then dProjLength = EPSILON

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = -dX
    dlinepts(7) = dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = -dX
    dlinepts(10) = dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    oAxisVec.Set 0, -1, 0
    Set ObjHvacBranchIns = PlaceProjectedBox(m_OutputColl, dlinepts, oAxisVec, dProjLength, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHvacBranchIns
    Set ObjHvacBranchIns = Nothing

    'New code for WCGOutputs
    Dim oGenBOFactory As IJGeneralBusinessObjectsFactory
    Dim oAttributes As IJDAttributes
    Dim oAttribCol As IJDAttributesCol
    Set oGenBOFactory = New GeneralBusinessObjectsFactory

    Set oAttributes = oGenBOFactory.CreateWCGPropHolder(m_OutputColl.ResourceManager)

    'In V8 we should use the IJWCGPropHolder interface to set these values
    Set oAttribCol = oAttributes.CollectionOfAttributes(CVar("IJWCGPropHolder"))

    Dim dInsVolume As Double
    Dim dHeaderVol As Double
    Dim dBranchVol As Double
    Dim dCGX As Double, dCGY As Double, dCGZ As Double

    'Calculate the Volume DI-CP·56531
    'Assumed that the thickness of duct is negligible
    dHeaderVol = 2 * (parWidth + 2 * parInsulationThickness + parDepth) * parHLength _
               * parInsulationThickness - _
                 ((parBWidth + 2 * parInsulationThickness) * (parBDepth + 2 * parInsulationThickness) _
                * parInsulationThickness)
    If parPlaneOfBranch = 0 Then
        dBranchVol = 2 * (parBWidth + 2 * parInsulationThickness + parBDepth) * _
                     (parBLength - parWidth / 2) * parInsulationThickness
    Else
        dBranchVol = 2 * (parBWidth + 2 * parInsulationThickness + parBDepth) * _
                     (parBLength - parDepth / 2) * parInsulationThickness
    End If
    dInsVolume = dHeaderVol + dBranchVol

    dCGX = 0    'as this symbol is symmetric about X-axis this is zero
    If Not CmpDblEqual(dInsVolume, LINEAR_TOLERANCE) Then
        If parPlaneOfBranch = 0 Then
            dCGY = (dBranchVol * (parWidth / 2 + (parBLength - parWidth / 2) / 2)) / dInsVolume
        Else
            dCGY = (dBranchVol * (parWidth / 2 + (parBLength - parDepth / 2) / 2)) / dInsVolume
        End If
    Else
        dCGY = 0
    End If
    dCGZ = 0    'as this symbol is symmetric about Z-axis this is zero

    oAttribCol.Item("Volume").Value = dInsVolume    '(cubic meters)
    oAttribCol.Item("CGX").Value = dCGX
    oAttribCol.Item("CGY").Value = dCGY
    oAttribCol.Item("CGZ").Value = dCGZ

    m_OutputColl.AddOutput "InsulationWCG", oAttributes
    Set oAttribCol = Nothing
    Set oAttributes = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                                   dPoints() As Double, _
                                   ByVal axisVector As AutoMath.DVector, _
                                   height As Double, _
                                   isCapped As Boolean) As Object

    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler

    Dim objProjection As IngrGeom3D.Projection3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d

    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = geomFactory.Projections3d.CreateByCurve( _
                        objOutputColl.ResourceManager, _
                        oLineStr, _
                        axisVector.x, axisVector.y, axisVector.z, _
                        height, isCapped)

    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Function
