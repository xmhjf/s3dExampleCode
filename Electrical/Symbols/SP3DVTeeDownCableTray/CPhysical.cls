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
'   Author:          svsmylav
'   Creation Date:  Friday, Mar 1 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is created using six outputs: 1,2 are objFrontTeeSection and objRearTeeSection- Created from
' the boundary curves, Output 3 is objCurvedSection is projected from curve, 4,5 and 6 are Cable Tray Ports.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.May.2004     MS     TR 58960 Removed hardcoded TangentLength value.
'                                           Now the TangentLength is retrieved from oTrayPart
'   23.Nov.2004     MS     CR 66680 Added input validation for TangentLength property for Tee and Cross
'                                           symbols (SP3DTeeCableTray, SP3DVTeeUpCableTray, SP3DVTeeDownCableTray, SP3DCrossCableTray)
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  13.Jun.2007     MA   CR-114880 Implemented Part data basis for values 71,73,75 and 77.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'   28.Jun.2011     GL      TR-199453 SP3DVTeeDownCableTray Symbol Def defines 2 outputs with the exact same name.
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
    
    Dim iOutput     As Double
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim parHeaderLength As Double
    Dim parTeeBranchLength As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parTangentLength3 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parHeaderLength = arrayOfInputs(2)
'    parTeeBranchLength = arrayOfInputs(3)
'    parTangentLength1 = arrayOfInputs(4)
'    parTangentLength2 = arrayOfInputs(5)
'    parTangentLength3 = arrayOfInputs(6)
    
    iOutput = 0
'   Retrieve Ports 1: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)

    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    Dim dTangentLength3 As Double   'Tangent length along Port 3.
    
' Insert your code for output 1: Tray Bottom
    Dim oPort1 As AutoMath.DPosition  'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition  'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    Dim oPort3 As AutoMath.DPosition  'Port 2 center point
    Set oPort3 = New AutoMath.DPosition
    Dim HD              As Double
    Dim HW              As Double
    Dim Port1S(0 To 11)  As Double
    Dim Port2S(0 To 11)  As Double
    Dim Port3S(0 To 11)  As Double
    Dim BFLA(0 To 8)  As Double
    Dim BFRA(0 To 8)  As Double

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -71):Vertical Curved Downward Three Way, Symmetrical, specified
'                               by Bend Radius and Tangent Length
'   (Part data Basis value -73):Vertical Curved Downward Three Way, Asymmetrical, specified
'                               by Bend Radius, Tangent Length 1, Tangent Length 2 and
'                               Tangent Length 3
'   (Part data Basis value -75):Vertical Segmented Downward Three Way, Symmetrical, specified
'                               by Header Length, Branch Length and Tangent Length
'   (Part data Basis value -77):Vertical Segmented Downward Three Way, Asymmetrical, specified
'                               by Header Length, Branch Length, Tangent Length 1,Tangent Length 2
'                               and Tangent Length 3

'Checking for the Part Data Basis Property
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    Dim lCableTrayPartDataBasis As Integer
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    
    Select Case lCableTrayPartDataBasis

        Case Is <= 1, 71
        '   Retrieve Part properties: BendRadius and TangentLength
            parBendRadius = oTrayPart.BendRadius
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
        
        '   Port 1 position
            oPort1.Set dTangentLength1 + parBendRadius + HW, 0, 0

        '   Port 2 position
            oPort2.Set -(dTangentLength2 + parBendRadius + HW), 0, 0

        '   Port 3 position
            oPort3.Set 0, 0, -(dTangentLength3 + parBendRadius + HW)
        
        Case 73
            parTangentLength1 = arrayOfInputs(4)
            parTangentLength2 = arrayOfInputs(5)
            parTangentLength3 = arrayOfInputs(6)
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dTangentLength3 = parTangentLength3
        
        '   Retrieve Part properties: BendRadius
            parBendRadius = oTrayPart.BendRadius
        
        '   Port 1 position
            oPort1.Set dTangentLength1 + parBendRadius + HW, 0, 0

        '   Port 2 position
            oPort2.Set -(dTangentLength2 + parBendRadius + HW), 0, 0

        '   Port 3 position
            oPort3.Set 0, 0, -(dTangentLength3 + parBendRadius + HW)
        
        Case 75
            parHeaderLength = arrayOfInputs(2)
            parTeeBranchLength = arrayOfInputs(3)
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
        
        '   Port 1 position
            oPort1.Set parHeaderLength / 2, 0, 0
        
        '   Port 2 position
            oPort2.Set -parHeaderLength / 2, 0, 0
        
        '   Port 3 position
            oPort3.Set 0, 0, -(parTeeBranchLength - HD)
            
        Case 77
            parHeaderLength = arrayOfInputs(2)
            parTeeBranchLength = arrayOfInputs(3)
            parTangentLength1 = arrayOfInputs(4)
            parTangentLength2 = arrayOfInputs(5)
            parTangentLength3 = arrayOfInputs(6)
        
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dTangentLength3 = parTangentLength3
        
        '   Port 1 position
            oPort1.Set parHeaderLength / 2, 0, 0
        
        '   Port 2 position
            oPort2.Set -parHeaderLength / 2, 0, 0
        
        '   Port 3 position
            oPort3.Set 0, 0, -(parTeeBranchLength - HD)
        
        Case Else
            GoTo ErrorLabel:

    End Select

'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength3, 0) Then dTangentLength3 = 0.0001

'   Tray Port 1 U-shape points positions
'   Front-top
    Port1S(0) = oPort1.x
    Port1S(1) = oPort1.y - HW
    Port1S(2) = oPort1.z + HD
'   Front-bottom
    Port1S(3) = oPort1.x
    Port1S(4) = oPort1.y - HW
    Port1S(5) = oPort1.z - HD
'   Back-bottom
    Port1S(6) = oPort1.x
    Port1S(7) = oPort1.y + HW
    Port1S(8) = oPort1.z - HD
'   Back-top
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y + HW
    Port1S(11) = oPort1.z + HD
    
'   Tray Port 2 points positions
'   Front-top
    Port2S(0) = oPort2.x
    Port2S(1) = oPort2.y - HW
    Port2S(2) = oPort2.z + HD
'   Front-bottom
    Port2S(3) = oPort2.x
    Port2S(4) = oPort2.y - HW
    Port2S(5) = oPort2.z - HD
'   Back-bottom
    Port2S(6) = oPort2.x
    Port2S(7) = oPort2.y + HW
    Port2S(8) = oPort2.z - HD
'   Back-top
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y + HW
    Port2S(11) = oPort2.z + HD
    
'   Tray Port 3 U-shape points positions
'   Front-Left
    Port3S(0) = oPort3.x - HD
    Port3S(1) = oPort3.y - HW
    Port3S(2) = oPort3.z
'   Front-Right
    Port3S(3) = oPort3.x + HD
    Port3S(4) = oPort3.y - HW
    Port3S(5) = oPort3.z
'   Back-right
    Port3S(6) = oPort3.x + HD
    Port3S(7) = oPort3.y + HW
    Port3S(8) = oPort3.z
'   Back-left
    Port3S(9) = oPort3.x - HD
    Port3S(10) = oPort3.y + HW
    Port3S(11) = oPort3.z
    
' Insert your code for output 1: Front Tee-Section
'   Branch Left Front Arc BFLA
    BFLA(0) = Port2S(3) + dTangentLength2
    BFLA(1) = Port2S(4)
    BFLA(2) = Port2S(5)

    BFLA(3) = Port3S(0)
    BFLA(4) = Port3S(1)
    BFLA(5) = Port3S(2) + dTangentLength3
    
    BFLA(6) = Port2S(3) + dTangentLength2
    BFLA(7) = Port3S(1)
    BFLA(8) = Port3S(2) + dTangentLength3

'   Construct Front Tee-Section: Header and Branch curves
'   Construct header curve
    Dim oHFLine           As IngrGeom3D.Line3d
    Set oHFLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(0), Port1S(1), Port1S(2), Port2S(0), Port2S(1), Port2S(2))
'   Construct Port2 Front Edge
    Dim oP2FEdge           As IngrGeom3D.Line3d
    Set oP2FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(0), Port2S(1), Port2S(2), Port2S(3), Port2S(4), Port2S(5))
'   Construct header front left tangent line
    Dim oFLTLine           As IngrGeom3D.Line3d
    Set oFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), BFLA(0), BFLA(1), BFLA(2))
'   Construct branch front left tangent line
    Dim oBFLTLine           As IngrGeom3D.Line3d
    Set oBFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(3), BFLA(4), BFLA(5), Port3S(0), Port3S(1), Port3S(2))
'   Construct Port3 Front Edge
    Dim oP3FEdge           As IngrGeom3D.Line3d
    Set oP3FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(0), Port3S(1), Port3S(2), Port3S(3), Port3S(4), Port3S(5))
'   Construct branch right tangent line
    Dim oBRTLine           As IngrGeom3D.Line3d
    Set oBRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), -BFLA(3), BFLA(4), BFLA(5))
'   Construct branch front right tangent line
    Dim oFRTLine           As IngrGeom3D.Line3d
    Set oFRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(0), BFLA(1), BFLA(2), Port1S(3), Port1S(4), Port1S(5))
'   Construct Port1 Front Edge
    Dim oP1FEdge           As IngrGeom3D.Line3d
    Set oP1FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), Port1S(0), Port1S(1), Port1S(2))
'   Construct entire Top header and branch curves
    Dim oFrontBranchcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
'   Dim objHcurves As IngrGeom3D.ComplexString3d
    Set oFrontBranchcurves = New Collection

'   Construct Rear Tee-Section: Header and Branch curves
'   Construct header curve
    Dim oHRLine           As IngrGeom3D.Line3d
    Set oHRLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(9), Port1S(10), Port1S(11), Port2S(9), Port2S(10), Port2S(11))
'   Construct Port2 Rear Edge
    Dim oP2REdge           As IngrGeom3D.Line3d
    Set oP2REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(9), Port2S(10), Port2S(11), Port2S(6), Port2S(7), Port2S(8))
'   Construct header Rear left tangent line
    Dim oRLTLine           As IngrGeom3D.Line3d
    Set oRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), BFLA(0), -BFLA(1), BFLA(2))
'   Construct branch Rear left tangent line
    Dim oBRLTLine           As IngrGeom3D.Line3d
    Set oBRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(3), -BFLA(4), BFLA(5), Port3S(9), Port3S(10), Port3S(11))
'   Construct Port3 Rear Edge
    Dim oP3REdge           As IngrGeom3D.Line3d
    Set oP3REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(9), Port3S(10), Port3S(11), Port3S(6), Port3S(7), Port3S(8))
'   Construct branch rear right tangent line
    Dim oBRRTLine           As IngrGeom3D.Line3d
    Set oBRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), -BFLA(3), -BFLA(4), BFLA(5))
'   Construct branch Rear right tangent line
    Dim oRRTLine           As IngrGeom3D.Line3d
    Set oRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(0), -BFLA(1), BFLA(2), Port1S(6), Port1S(7), Port1S(8))
'   Construct Port1 Rear Edge
    Dim oP1REdge           As IngrGeom3D.Line3d
    Set oP1REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), Port1S(9), Port1S(10), Port1S(11))
'   Construct entire Bottom header and branch curves
    Dim oRearBranchcurves           As Collection
    Dim objRearHBcurves         As IngrGeom3D.ComplexString3d
    Dim objRearHcurves As IngrGeom3D.ComplexString3d
    Set oRearBranchcurves = New Collection
    
    Dim objBp1scurve         As IngrGeom3D.ComplexString3d
    Dim oBp1scurve           As Collection
    Set oBp1scurve = New Collection
    
    Dim objBp1scurve1         As IngrGeom3D.ComplexString3d
    Dim oBp1scurve1           As Collection
    Set oBp1scurve1 = New Collection

'Checking for the Part Data Basis Property
    Set oTrayPart = oPartFclt
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    
    Select Case lCableTrayPartDataBasis

        Case Is <= 1, 71, 73
        '   Construct branch front left curve
            Dim objBFLA  As IngrGeom3D.Arc3d
            Dim SBFLA   As AutoMath.DPosition
            Set SBFLA = New AutoMath.DPosition
            Dim EBFLA   As AutoMath.DPosition
            Set EBFLA = New AutoMath.DPosition
            Dim CBFLA   As AutoMath.DPosition
            Set CBFLA = New AutoMath.DPosition
            SBFLA.Set BFLA(0), BFLA(1), BFLA(2)
            EBFLA.Set BFLA(3), BFLA(4), BFLA(5)
            CBFLA.Set BFLA(6), BFLA(7), BFLA(8)
            Set objBFLA = PlaceTrArcByCenter(SBFLA, EBFLA, CBFLA)

        '   Construct branch front right curve
            Dim objBFRA  As IngrGeom3D.Arc3d
            Dim SBFRA   As AutoMath.DPosition
            Set SBFRA = New AutoMath.DPosition
            Dim EBFRA   As AutoMath.DPosition
            Set EBFRA = New AutoMath.DPosition
            Dim CBFRA   As AutoMath.DPosition
            Set CBFRA = New AutoMath.DPosition
            SBFRA.Set -BFLA(3), BFLA(4), BFLA(5)
            EBFRA.Set -BFLA(0), BFLA(1), BFLA(2)
            CBFRA.Set -BFLA(6), BFLA(7), BFLA(8)
            Set objBFRA = PlaceTrArcByCenter(SBFRA, EBFRA, CBFRA)

            oFrontBranchcurves.Add oHFLine
            oFrontBranchcurves.Add oP2FEdge
            oFrontBranchcurves.Add oFLTLine
            oFrontBranchcurves.Add objBFLA
            oFrontBranchcurves.Add oBFLTLine
            oFrontBranchcurves.Add oP3FEdge
            oFrontBranchcurves.Add oBRTLine
            oFrontBranchcurves.Add objBFRA
            oFrontBranchcurves.Add oFRTLine
            oFrontBranchcurves.Add oP1FEdge
            
        '   Construct branch Rear left curve Edge
            Dim objBRLA  As IngrGeom3D.Arc3d
            Dim SBRLA   As AutoMath.DPosition
            Set SBRLA = New AutoMath.DPosition
            Dim EBRLA   As AutoMath.DPosition
            Set EBRLA = New AutoMath.DPosition
            Dim CBRLA   As AutoMath.DPosition
            Set CBRLA = New AutoMath.DPosition
            SBRLA.Set BFLA(0), -BFLA(1), BFLA(2)
            EBRLA.Set BFLA(3), -BFLA(4), BFLA(5)
            CBRLA.Set BFLA(6), -BFLA(7), BFLA(8)
            Set objBRLA = PlaceTrArcByCenter(SBRLA, EBRLA, CBRLA)
            
        '   Construct branch Rear right curve
            Dim objBRRA  As IngrGeom3D.Arc3d
            Dim SBRRA   As AutoMath.DPosition
            Set SBRRA = New AutoMath.DPosition
            Dim EBRRA   As AutoMath.DPosition
            Set EBRRA = New AutoMath.DPosition
            Dim CBRRA   As AutoMath.DPosition
            Set CBRRA = New AutoMath.DPosition
            SBRRA.Set -BFLA(3), -BFLA(4), BFLA(5)
            EBRRA.Set -BFLA(0), -BFLA(1), BFLA(2)
            CBRRA.Set -BFLA(6), -BFLA(7), BFLA(8)
            Set objBRRA = PlaceTrArcByCenter(SBRRA, EBRRA, CBRRA)

            oRearBranchcurves.Add oHRLine
            oRearBranchcurves.Add oP2REdge
            oRearBranchcurves.Add oRLTLine
            oRearBranchcurves.Add objBRLA
            oRearBranchcurves.Add oBRLTLine
            oRearBranchcurves.Add oP3REdge
            oRearBranchcurves.Add oBRRTLine
            oRearBranchcurves.Add objBRRA
            oRearBranchcurves.Add oRRTLine
            oRearBranchcurves.Add oP1REdge
            
            oBp1scurve.Add oBRTLine
            oBp1scurve.Add objBFRA
            oBp1scurve.Add oFRTLine
            
            oBp1scurve1.Add oFLTLine
            oBp1scurve1.Add objBFLA
            oBp1scurve1.Add oBFLTLine

        Case 75, 77
        '   Construct branch front left line
            Dim oBFLL           As IngrGeom3D.Line3d
            Set oBFLL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(0), BFLA(1), BFLA(2), BFLA(3), BFLA(4), BFLA(5))

        '   Construct branch front right line
            Dim oBFRL           As IngrGeom3D.Line3d
            Set oBFRL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(3), BFLA(4), BFLA(5), -BFLA(0), BFLA(1), BFLA(2))
            
            oFrontBranchcurves.Add oHFLine
            oFrontBranchcurves.Add oP2FEdge
            oFrontBranchcurves.Add oFLTLine
            oFrontBranchcurves.Add oBFLL
            oFrontBranchcurves.Add oBFLTLine
            oFrontBranchcurves.Add oP3FEdge
            oFrontBranchcurves.Add oBRTLine
            oFrontBranchcurves.Add oBFRL
            oFrontBranchcurves.Add oFRTLine
            oFrontBranchcurves.Add oP1FEdge
            
        '   Construct branch Rear left line
            Dim oBRLL           As IngrGeom3D.Line3d
            Set oBRLL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(0), -BFLA(1), BFLA(2), BFLA(3), -BFLA(4), BFLA(5))
            
        '   Construct branch Rear right line
            Dim oBRRL           As IngrGeom3D.Line3d
            Set oBRRL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(3), -BFLA(4), BFLA(5), -BFLA(0), -BFLA(1), BFLA(2))
            
            oRearBranchcurves.Add oHRLine
            oRearBranchcurves.Add oP2REdge
            oRearBranchcurves.Add oRLTLine
            oRearBranchcurves.Add oBRLL
            oRearBranchcurves.Add oBRLTLine
            oRearBranchcurves.Add oP3REdge
            oRearBranchcurves.Add oBRRTLine
            oRearBranchcurves.Add oBRRL
            oRearBranchcurves.Add oRRTLine
            oRearBranchcurves.Add oP1REdge
            
            oBp1scurve.Add oBRTLine
            oBp1scurve.Add oBFRL
            oBp1scurve.Add oFRTLine
            
            oBp1scurve1.Add oFLTLine
            oBp1scurve1.Add oBFLL
            oBp1scurve1.Add oBFLTLine
            
        Case Else
            GoTo ErrorLabel:

    End Select
    
'   Construct front-tee
    Dim StartBC   As AutoMath.DPosition
    Set StartBC = New AutoMath.DPosition
    StartBC.Set Port1S(0), Port1S(1), Port1S(2)
    Set objHBcurves = PlaceTrCString(StartBC, oFrontBranchcurves)
    Dim ObjFrontTeeSection As IngrGeom3D.Plane3d
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 1, 0
    Set ObjFrontTeeSection = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjFrontTeeSection.AddBoundary(objHBcurves)
    'Remove cable tray bottom Header and Branch lines
    Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontTeeSection
    Set ObjFrontTeeSection = Nothing
    Set StartBC = Nothing

'   Construct rear-tee
    Dim StartBCRear   As AutoMath.DPosition
    Set StartBCRear = New AutoMath.DPosition
    StartBCRear.Set Port1S(9), Port1S(10), Port1S(11)
    Set objRearHBcurves = PlaceTrCString(StartBCRear, oRearBranchcurves)
    Dim ObjRearTeeSection As IngrGeom3D.Plane3d
    Dim oDirProjRear        As AutoMath.DVector
    Set oDirProjRear = New AutoMath.DVector
    oDirProjRear.Set 0, 1, 0
    Set ObjRearTeeSection = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(6), Port1S(7), Port1S(8), oDirProjRear.x, oDirProjRear.y, oDirProjRear.z)
    Call ObjRearTeeSection.AddBoundary(objRearHBcurves)
    'Remove cable tray bottom Header and Branch lines
'   Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objRearHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearTeeSection
    Set ObjRearTeeSection = Nothing
    Set oDirProjRear = Nothing
    Set StartBCRear = Nothing

' Insert your code for output 3: BranchPort1Side
    Dim StartBCP3   As AutoMath.DPosition
    Set StartBCP3 = New AutoMath.DPosition
    StartBCP3.Set Port3S(3), Port3S(4), Port3S(5)
    
    Set objBp1scurve = PlaceTrCString(StartBCP3, oBp1scurve)
    Dim objCurvedSection As IJDObject
    Set objCurvedSection = PlaceProjection(m_OutputColl, objBp1scurve, oDirProj, parActualWidth, True)
    Set objBp1scurve = Nothing

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSection
    Set StartBCP3 = Nothing

' Insert your code for output: BranchPort2Side
    Dim StartBCP4   As AutoMath.DPosition
    Set StartBCP4 = New AutoMath.DPosition
    StartBCP4.Set Port2S(3), Port2S(4), Port2S(5)
    
    Set objBp1scurve1 = PlaceTrCString(StartBCP4, oBp1scurve1)
    Set objCurvedSection = PlaceProjection(m_OutputColl, objBp1scurve1, oDirProj, parActualWidth, True)
    
    m_OutputColl.AddOutput "CurvedSection2", objCurvedSection
    Set objBp1scurve1 = Nothing
    Set oDirProj = Nothing
    Set objCurvedSection = Nothing
    Set StartBCP4 = Nothing

'   Remove branch curves
    Set objBFLA = Nothing
    Set objBFRA = Nothing
    Set objBRLA = Nothing
    Set objBRRA = Nothing
    Set oHFLine = Nothing
    Set oP1FEdge = Nothing
    Set oP2FEdge = Nothing
    Set oP3FEdge = Nothing
    Set oP1REdge = Nothing
    Set oP2REdge = Nothing
    Set oP3REdge = Nothing
    Set oFLTLine = Nothing
    Set oBFLTLine = Nothing
    Set oRLTLine = Nothing
    Set oBRTLine = Nothing
    Set oFRTLine = Nothing
    Set oRRTLine = Nothing
    Set oHRLine = Nothing
    Set oBRLTLine = Nothing
    Set oBRRTLine = Nothing
    Set oBFRL = Nothing
    Set oBFLL = Nothing
    Set oBRRL = Nothing
    Set oBRLL = Nothing

    Dim iCount As Integer
    For iCount = 1 To oFrontBranchcurves.Count
        oFrontBranchcurves.Remove 1
    Next iCount
    Set oFrontBranchcurves = Nothing
    
    For iCount = 1 To oRearBranchcurves.Count
        oRearBranchcurves.Remove 1
    Next iCount
    Set oRearBranchcurves = Nothing
    
    For iCount = 1 To oBp1scurve.Count
        oBp1scurve.Remove 1
    Next iCount
    Set oBp1scurve = Nothing
    
' Place Port 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

    oDir.Set 1, 0, 0
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

    oDir.Set -1, 0, 0
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

    oDir.Set 0, 0, -1
    oRadialOrient.Set -1, 0, 0

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
