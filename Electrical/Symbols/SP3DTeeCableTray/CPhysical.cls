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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Saturday, Feb 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy         who                     change description
'   -----------        -----                    ------------------
'   09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'   24.Oct.2003     SymbolTeam(India)  TR 51102 Removed duplicated Header/Copyright Information.
'   23.Nov.2004     MS                 CR 66680 Added input validation for TangentLength property for Tee and Cross
'                                      symbols (SP3DTeeCableTray, SP3DVTeeUpCableTray, SP3DVTeeDownCableTray, SP3DCrossCableTray)
'   08.SEP.2006     KKC                DI-95670 Replace names with initials in all revision history sheets and symbols
'   11.June.2007    KKC                DM-121367 Implemented part data basis for the Square Outside Corners.
'   14.JUN.2007     dkl                CR-114880 Implemented part data basis for segmented Tee. (Refer Horizontal Three Way, Zhen Hua Catalog.)
'   3.JUL.2007      dkl                CR-122171 Implemented part data bases 111 for Schneider Electric Catalog.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
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
    
    Dim iOutput As Double
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim parActualWidth3 As Double
    Dim lCableTrayPartDataBasis As Integer
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parCTDimension = arrayOfInputs(2)
'    parHeaderLength = arrayOfInputs(3)
'    parTangentLength1 = arrayOfInputs(4)
'    parTangentLength2 = arrayOfInputs(5)
'    parTangentLength3 = arrayOfInputs(6)
'    parTeeBranchLength = arrayOfInputs(7)

    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    Dim oPort3 As AutoMath.DPosition 'Port 3 center point
    Set oPort3 = New AutoMath.DPosition
    
    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    Dim dTangentLength3 As Double   'Tangent length along Port 3.
    
    iOutput = 0
   
'   Retrieve Ports 1 & 3 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Call RetrieveCableTrayPortProperties(3, oPartFclt, parActualWidth3, parActualDepth)
    
    Dim HD              As Double
    Dim HW              As Double
    Dim HW3              As Double
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    HW3 = parActualWidth3 / 2
    
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
'   The following Part data Bases are implemented,
'   1) Curved Horizontal Three Way, Symmetrical, specified by Bend Radius and Tangent Length (Default implementation).
'   2) Bend Radius Dimension Exclusion basis .
'   3) Curved Horizontal Three Way, Asymmetrical, specified by Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
'   4) Curved Horizontal Three Way, Symmetrical, specified by Header Length, Bend Radius and Tangent Length.
'   5) Curved Horizontal Three Way, Asymmetrical, specified by Header Length, Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
'   6) Segmented Horizontal Three Way, Symmetrical, specified by Header Length, Bend Radius and Tangent Length.
'   7) Segmented Horizontal Three Way, Asymmetrical, specified by Header Length, Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
'   8) Curved Horizontal Three Way, specified by Header Length, Tee Branch Length and Bend Radius.

    Select Case lCableTrayPartDataBasis
        Case Is <= 1, 7     '7 - Curved Horizontal Three Way, Symmetrical, specified by
                            'Bend Radius and Tangent Length (Default implementation).
            parTangentLength = oTrayPart.TangentLength
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            
            oPort1.Set -(dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort2.Set (dTangentLength2 + parBendRadius + HW3), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            
        Case 5      'Bend Radius Dimension Exclusion basis.
            Dim parCTDimension As Double
            parCTDimension = arrayOfInputs(2)
            parTangentLength = oTrayPart.TangentLength
            
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            
            oPort1.Set -(parCTDimension + HW3), 0, 0
            oPort2.Set (parCTDimension + HW3), 0, 0
            oPort3.Set 0, (parCTDimension + HW), 0
            
        Case 9          'Curved Horizontal Three Way, Asymmetrical, specified by
                        'Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
            dTangentLength1 = arrayOfInputs(4)
            dTangentLength2 = arrayOfInputs(5)
            dTangentLength3 = arrayOfInputs(6)
            parBendRadius = oTrayPart.BendRadius
        
            oPort1.Set -(dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort2.Set (dTangentLength2 + parBendRadius + HW3), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            
        Case 11, 15     '11 - Curved Horizontal Three Way, Symmetrical, specified by
                        ' Header Length, Bend Radius and Tangent Length.
                        '15 - Segmented Horizontal Three Way, Symmetrical, specified by
                        'Header Length, Bend Radius and Tangent Length.
            Dim parHeaderLength As Double
            parHeaderLength = arrayOfInputs(3)
            parTangentLength = oTrayPart.TangentLength
            parBendRadius = oTrayPart.BendRadius
            
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            
            oPort1.Set -(dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort2.Set parHeaderLength - (dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            
        Case 13, 17     '13 - Curved Horizontal Three Way, Asymmetrical, specified by Header Length,
                        'Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
                        '17 - Segmented Horizontal Three Way, Asymmetrical, specified by Header Length,
                        ' Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
            parHeaderLength = arrayOfInputs(3)
            dTangentLength1 = arrayOfInputs(4)
            dTangentLength2 = arrayOfInputs(5)
            dTangentLength3 = arrayOfInputs(6)
            parBendRadius = oTrayPart.BendRadius
                        
            oPort1.Set -(dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort2.Set parHeaderLength - (dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
        
        Case 111    'Curved Horizontal Three Way, specified by
                    'Header Length, Tee Branch Length and Bend Radius
            Dim parTeeBranchLength As Double
            parHeaderLength = arrayOfInputs(3)
            parTeeBranchLength = arrayOfInputs(7)
            parBendRadius = oTrayPart.BendRadius
            
            dTangentLength1 = (parHeaderLength - parActualWidth3 - 2 * parBendRadius) / 2
            dTangentLength2 = dTangentLength1
            dTangentLength3 = parTeeBranchLength - parActualWidth - parBendRadius
            
            oPort1.Set -(dTangentLength1 + parBendRadius + HW3), 0, 0
            oPort2.Set dTangentLength2 + parBendRadius + HW3, 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            
        Case Else
            GoTo ErrorLabel:
    
    End Select

    If CmpDblLessThan(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThan(dTangentLength2, 0) Then dTangentLength2 = 0.0001
    If CmpDblLessThan(dTangentLength3, 0) Then dTangentLength3 = 0.0001

 ' Insert your code for output 1: Tray Bottom

    Dim Port1S(0 To 11)  As Double
    Dim Port2S(0 To 11)  As Double
    Dim Port3S(0 To 11)  As Double
    Dim BLA(0 To 8)  As Double
    Dim BRA(0 To 8)  As Double
    Dim BHC(0 To 17)  As Double
    Dim HSide(0 To 11)  As Double
    
'   Port 1 position
    
'   Tray Port 1 U-shape points positions

'   Tray top edge close to curved branch
    Port1S(0) = oPort1.x
    Port1S(1) = oPort1.y + HW
    Port1S(2) = oPort1.z + HD
'   Tray bottom
    Port1S(3) = oPort1.x
    Port1S(4) = oPort1.y + HW
    Port1S(5) = oPort1.z - HD
    
    Port1S(6) = oPort1.x
    Port1S(7) = oPort1.y - HW
    Port1S(8) = oPort1.z - HD
'   Tray top edge far from curved branch
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD
    
'   Port 2 position

'   Tray Port 2 U-shape points positions
'   Tray top edge close to curved branch
    Port2S(0) = oPort2.x
    Port2S(1) = oPort2.y + HW
    Port2S(2) = oPort2.z + HD
'   Tray bottom
    Port2S(3) = oPort2.x
    Port2S(4) = oPort2.y + HW
    Port2S(5) = oPort2.z - HD
    
    Port2S(6) = oPort2.x
    Port2S(7) = oPort2.y - HW
    Port2S(8) = oPort2.z - HD
'   Tray top edge far from curved branch
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD
    
'   Port 3 position
    
'   Tray Port 3 U-shape points positions
'   Tray top edge close to curved branch
    Port3S(0) = oPort3.x - HW3
    Port3S(1) = oPort3.y
    Port3S(2) = oPort3.z + HD
'   Tray bottom
    Port3S(3) = oPort3.x - HW3
    Port3S(4) = oPort3.y
    Port3S(5) = oPort3.z - HD
    
    Port3S(6) = oPort3.x + HW3
    Port3S(7) = oPort3.y
    Port3S(8) = oPort3.z - HD
'   Tray top edge far from curved branch
    Port3S(9) = oPort3.x + HW3
    Port3S(10) = oPort3.y
    Port3S(11) = oPort3.z + HD
    
'   Branch Left Arc BLA
    BLA(0) = Port1S(3) + dTangentLength1
    BLA(1) = Port1S(4)
    BLA(2) = Port1S(5)
'   Tray bottom
    BLA(3) = Port3S(3)
    BLA(4) = Port3S(4) - dTangentLength3
    BLA(5) = Port3S(5)
    
    BLA(6) = Port1S(3) + dTangentLength1
    BLA(7) = Port3S(4) - dTangentLength3
    BLA(8) = Port1S(5)

'   Construct Tray bottom: Header and Branch curves
'   Construct bottom header curve
    Dim oHLine           As IngrGeom3D.Line3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), Port1S(6), Port1S(7), Port1S(8))
'   Construct Branch curve
    Dim oBLine           As IngrGeom3D.Line3d
    Set oBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(6), Port3S(7), Port3S(8))
'   Construct Header Left Tangent line
    Dim oLTLine           As IngrGeom3D.Line3d
    Set oLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), BLA(0), BLA(1), BLA(2))
'   Construct branch left tangent line
    Dim oBLTLine           As IngrGeom3D.Line3d
    Set oBLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLA(3), BLA(4), BLA(5), Port3S(3), Port3S(4), Port3S(5))
'   Construct branch right tangent line
    Dim oBRTLine           As IngrGeom3D.Line3d
    Set oBRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), -BLA(3), BLA(4), BLA(5))
'   Construct Header Right Tangent line
    Dim oRTLine           As IngrGeom3D.Line3d
    Set oRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BLA(0), BLA(1), BLA(2), Port2S(3), Port2S(4), Port2S(5))
'   Construct Port1 bottom curve
    Dim oPort1botLine           As IngrGeom3D.Line3d
    Set oPort1botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), Port1S(3), Port1S(4), Port1S(5))
'   Construct Port2 bottom curve
    Dim oPort2botLine           As IngrGeom3D.Line3d
    Set oPort2botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), Port2S(6), Port2S(7), Port2S(8))

'   Construct collection for creating the bottom surface.
    Dim obranchcurves As Collection
    Set obranchcurves = New Collection
    'Adding curves for the Tray Bottom.
    obranchcurves.Add oLTLine

'   Construct collection for creating Branch Port 1 Side.
    Dim oBp1scurve As Collection
    Set oBp1scurve = New Collection
    'Adding curves for the Branch Port 1 Side.
    oBp1scurve.Add oLTLine
    
'   Construct collection for creating Branch Port 2 Side.
    Dim oBp2scurve As Collection
    Set oBp2scurve = New Collection
    'Adding curves for the Branch Port 2 Side.
    oBp2scurve.Add oBRTLine
    
    Select Case lCableTrayPartDataBasis
        Case Is <= 1, 7, 9, 11, 13, 111
            ' Construct branch left curve
            Dim objBLA  As IngrGeom3D.Arc3d
            Dim SBLA   As New AutoMath.DPosition
            Dim EBLA   As New AutoMath.DPosition
            Dim CBLA   As New AutoMath.DPosition
            SBLA.Set BLA(0), BLA(1), BLA(2)
            EBLA.Set BLA(3), BLA(4), BLA(5)
            CBLA.Set BLA(6), BLA(7), BLA(8)
            Set objBLA = PlaceTrArcByCenter(SBLA, EBLA, CBLA)
            
            ' Construct branch right curve
            Dim objBRA  As IngrGeom3D.Arc3d
            Dim SBRA   As New AutoMath.DPosition
            Dim EBRA   As New AutoMath.DPosition
            Dim CBRA   As New AutoMath.DPosition
            SBRA.Set -BLA(3), BLA(4), BLA(5)
            EBRA.Set -BLA(0), BLA(1), BLA(2)
            CBRA.Set -BLA(6), BLA(7), BLA(8)
            Set objBRA = PlaceTrArcByCenter(SBRA, EBRA, CBRA)
            
            'Adding curves for the Tray Bottom.
            obranchcurves.Add objBLA
            obranchcurves.Add oBLTLine
            obranchcurves.Add oBLine
            obranchcurves.Add oBRTLine
            obranchcurves.Add objBRA
            
            'Adding curves for the Branch Port 1 Side.
            oBp1scurve.Add objBLA
            
            'Adding curves for the Branch Port 2 Side.
            oBp2scurve.Add objBRA

        Case 5, 15, 17
            'Construct Left Inclined line.
            Dim objLTL  As IngrGeom3D.Line3d
            Set objLTL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLA(0), BLA(1), BLA(2), BLA(3), BLA(4), BLA(5))
            
            'Construct Right inclined line.
            Dim objRTL  As IngrGeom3D.Line3d
            Set objRTL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BLA(3), BLA(4), BLA(5), -BLA(0), BLA(1), BLA(2))
            
            'Adding curves for the Tray Bottom.
            obranchcurves.Add objLTL
            obranchcurves.Add oBLTLine
            obranchcurves.Add oBLine
            obranchcurves.Add oBRTLine
            obranchcurves.Add objRTL
            
            'Adding curves for the Branch Port 1 Side.
            oBp1scurve.Add objLTL
            
            'Adding curves for the Branch Port 2 Side.
            oBp2scurve.Add objRTL
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    'Adding curves for the Tray Bottom.
    obranchcurves.Add oRTLine
    obranchcurves.Add oPort2botLine
    obranchcurves.Add oHLine
    obranchcurves.Add oPort1botLine

    'Adding curves for the Branch Port 1 Side.
    oBp1scurve.Add oBLTLine
    
    'Adding curves for the Branch Port 2 Side.
    oBp2scurve.Add oRTLine
    
    Set oLTLine = Nothing
    Set objBLA = Nothing
    Set oBLTLine = Nothing
    Set oBLine = Nothing
    Set oBRTLine = Nothing
    Set objBRA = Nothing
    Set oRTLine = Nothing
    Set oPort2botLine = Nothing
    Set oHLine = Nothing
    Set oPort1botLine = Nothing
    
'Insert code for creating bottom surface.
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    
    Dim objHBcurves As IngrGeom3D.ComplexString3d
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    
    Dim TrayBottom As IngrGeom3D.Plane3d
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    Set TrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call TrayBottom.AddBoundary(objHBcurves)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), TrayBottom
    Set TrayBottom = Nothing
    Set objHBcurves = Nothing
    Set obranchcurves = Nothing
    
' Insert your code for output 2: Header side
    Dim HeaderSide As IngrGeom3D.Plane3d
    HSide(0) = Port1S(6)
    HSide(1) = Port1S(7)
    HSide(2) = Port1S(8)
    
    HSide(3) = Port2S(6)
    HSide(4) = Port2S(7)
    HSide(5) = Port2S(8)
    
    HSide(6) = Port2S(9)
    HSide(7) = Port2S(10)
    HSide(8) = Port2S(11)
    
    HSide(9) = Port1S(9)
    HSide(10) = Port1S(10)
    HSide(11) = Port1S(11)
    
    Set HeaderSide = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, HSide)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), HeaderSide
    Set HeaderSide = Nothing
    Set oGeomFactory = Nothing
    
' Insert your code for output 3: BranchPort1Side
    Dim objBp1scurve As IngrGeom3D.ComplexString3d
    Set objBp1scurve = PlaceTrCString(StartBC, oBp1scurve)
    
    Dim ObjBranchPort1Side As IJDObject
    Set ObjBranchPort1Side = PlaceProjection(m_OutputColl, objBp1scurve, oDirProj, _
                                                                        parActualDepth, True)
       
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchPort1Side
    Set ObjBranchPort1Side = Nothing
    Set objBp1scurve = Nothing
    Set oBp1scurve = Nothing
    
' Insert your code for output 4: BranchPort2Side
    Dim StartBC2   As New AutoMath.DPosition
    StartBC2.Set Port3S(6), Port3S(7), Port3S(8)
    
    Dim objBp2scurve         As IngrGeom3D.ComplexString3d
    Set objBp2scurve = PlaceTrCString(StartBC2, oBp2scurve)
    
    Dim ObjBranchPort2Side As IJDObject
    Set ObjBranchPort2Side = PlaceProjection(m_OutputColl, objBp2scurve, oDirProj, parActualDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchPort2Side
    Set ObjBranchPort2Side = Nothing
    Set objBp2scurve = Nothing
    Set oBp2scurve = Nothing
    Set oDirProj = Nothing
    
' Place Port 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc

    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Port 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1

    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Port 3
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1

    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort3 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
