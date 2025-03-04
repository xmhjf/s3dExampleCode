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
'   Creation Date:  Thursday, Mar 14 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is created using six outputs: Output 1 is ObjTrayBottom is 3D plane, 2 and 3 are
' objFrontTeeSection and objRearTeeSection- Created from the boundary curves, is projected from curve,
' 4,5 and 6 are Cable Tray Ports.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Nov.2004     MS     CR 66680 Added input validation for TangentLength property for Tee and Cross
'                                           symbols (SP3DTeeCableTray, SP3DVTeeUpCableTray, SP3DVTeeDownCableTray, SP3DCrossCableTray)
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  11.Jun.2007     MA   CR-114880 Implemented Part data basis for values 63,65,67 and 69.
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
    
    Dim iOutput     As Double
    
    Dim ActualWidth As Double
    Dim ActualDepth As Double
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
    Call RetrieveCableTrayPortProperties(1, oPartFclt, ActualWidth, ActualDepth)

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
    Dim HD                      As Double
    Dim HW                      As Double
    Dim Port1S(0 To 11)         As Double
    Dim Port2S(0 To 11)         As Double
    Dim Port3S(0 To 11)         As Double
    Dim BFLA(0 To 8)            As Double
    Dim BottomPlane(0 To 11)    As Double

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    HD = ActualDepth / 2
    HW = ActualWidth / 2
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -63):Vertical Curved Upward Three Way, Symmetrical, specified by
'                               Bend Radius and Tangent Length
'   (Part data Basis value -65):Vertical Curved Upward Three Way, Asymmetrical, specified by
'                               Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3
'   (Part data Basis value -67):Vertical Segmented Upward Three Way, Symmetrical, specified
'                               by Header Length, Branch Length and Tangent Length
'   (Part data Basis value -69):Vertical Segmented Upward Three Way, Asymmetrical, specified
'                               by Header Length, Branch Length, Tangent Length 1, Tangent Length 2
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

        Case Is <= 1, 63
        '   Retrieve Part properties: BendRadius and TangentLength
            parBendRadius = oTrayPart.BendRadius
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
        
        '   Port 1 position
            oPort1.Set -(dTangentLength1 + parBendRadius + HD), 0, 0

        '   Port 2 position
            oPort2.Set dTangentLength2 + parBendRadius + HD, 0, 0

        '   Port 3 position
            oPort3.Set 0, 0, dTangentLength3 + parBendRadius + HD
        
        Case 65
            parTangentLength1 = arrayOfInputs(4)
            parTangentLength2 = arrayOfInputs(5)
            parTangentLength3 = arrayOfInputs(6)
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dTangentLength3 = parTangentLength3
        
        '   Retrieve Part properties: BendRadius
            parBendRadius = oTrayPart.BendRadius
        
        '   Port 1 position
            oPort1.Set -(dTangentLength1 + parBendRadius + HD), 0, 0

        '   Port 2 position
            oPort2.Set dTangentLength2 + parBendRadius + HD, 0, 0

        '   Port 3 position
            oPort3.Set 0, 0, dTangentLength3 + parBendRadius + HD
        
        Case 67
            parHeaderLength = arrayOfInputs(2)
            parTeeBranchLength = arrayOfInputs(3)
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            
        '   Port 1 position
            oPort1.Set -parHeaderLength / 2, 0, 0
            
        '   Port 2 position
            oPort2.Set parHeaderLength / 2, 0, 0
                    
        '   Port 3 position
            oPort3.Set 0, 0, parTeeBranchLength - HD
            
        Case 69
            parHeaderLength = arrayOfInputs(2)
            parTeeBranchLength = arrayOfInputs(3)
            parTangentLength1 = arrayOfInputs(4)
            parTangentLength2 = arrayOfInputs(5)
            parTangentLength3 = arrayOfInputs(6)
            
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dTangentLength3 = parTangentLength3

        '   Port 1 position
            oPort1.Set -parHeaderLength / 2, 0, 0
            
        '   Port 2 position
            oPort2.Set parHeaderLength / 2, 0, 0
                    
        '   Port 3 position
            oPort3.Set 0, 0, parTeeBranchLength - HD
            
            
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

'   Tray bottom plane points
    BottomPlane(0) = Port1S(3)
    BottomPlane(1) = Port1S(4)
    BottomPlane(2) = Port1S(5)
    
    BottomPlane(3) = Port2S(3)
    BottomPlane(4) = Port2S(4)
    BottomPlane(5) = Port2S(5)
    
    BottomPlane(6) = Port2S(6)
    BottomPlane(7) = Port2S(7)
    BottomPlane(8) = Port2S(8)
    
    BottomPlane(9) = Port1S(6)
    BottomPlane(10) = Port1S(7)
    BottomPlane(11) = Port1S(8)
    
    Dim ObjTrayBottom As IngrGeom3D.Plane3d
    Set ObjTrayBottom = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, BottomPlane)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrayBottom
    Set ObjTrayBottom = Nothing

' Insert your code for output 2: Front Tee-Section
'   Branch Front Left Arc BFLA
    BFLA(0) = Port1S(0) + dTangentLength1
    BFLA(1) = Port1S(1)
    BFLA(2) = Port1S(2)

    BFLA(3) = Port3S(0)
    BFLA(4) = Port3S(1)
    BFLA(5) = Port3S(2) - dTangentLength3
    
    BFLA(6) = Port1S(0) + dTangentLength1
    BFLA(7) = Port3S(1)
    BFLA(8) = Port3S(2) - dTangentLength3

'   Construct Front Tee-Section: Header and Branch curves
'   Construct header curve
    Dim oHFLine           As IngrGeom3D.Line3d
    Set oHFLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), Port2S(3), Port2S(4), Port2S(5))
'   Construct Port2 Front Edge
    Dim oP2FEdge           As IngrGeom3D.Line3d
    Set oP2FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), Port2S(0), Port2S(1), Port2S(2))
'   Construct front right tangent line till the branch arc start point
    Dim oFRTLine           As IngrGeom3D.Line3d
    Set oFRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(0), Port2S(1), Port2S(2), -BFLA(0), BFLA(1), BFLA(2))
'   Construct branch front right tangent line starting from branch arc end point
    Dim oBFRTLine           As IngrGeom3D.Line3d
    Set oBFRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(3), BFLA(4), BFLA(5), Port3S(3), Port3S(4), Port3S(5))
'   Construct Port3 Front Edge
    Dim oP3FEdge           As IngrGeom3D.Line3d
    Set oP3FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(0), Port3S(1), Port3S(2))
'   Construct branch front left tangent line
    Dim oBFLTLine           As IngrGeom3D.Line3d
    Set oBFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(0), Port3S(1), Port3S(2), BFLA(3), BFLA(4), BFLA(5))
'   Construct front left tangent line from the branch arc start point
    Dim oFLTLine           As IngrGeom3D.Line3d
    Set oFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(0), BFLA(1), BFLA(2), Port1S(0), Port1S(1), Port1S(2))
'   Construct Port1 Front Edge
    Dim oP1FEdge           As IngrGeom3D.Line3d
    Set oP1FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(0), Port1S(1), Port1S(2), Port1S(3), Port1S(4), Port1S(5))

'   Construct entire Top header and branch curves
    Dim oFrontBranchcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Set oFrontBranchcurves = New Collection

' Insert your code for output 3: Rear Tee-Section
'   Construct Rear Tee-Section: Header and Branch curves
'   Construct header curve
    Dim oHRLine           As IngrGeom3D.Line3d
    Set oHRLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), Port2S(6), Port2S(7), Port2S(8))
'   Construct Port2 Rear Edge
    Dim oP2REdge           As IngrGeom3D.Line3d
    Set oP2REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), Port2S(9), Port2S(10), Port2S(11))
'   Construct Rear right tangent line till the branch arc start point
    Dim oRRTLine           As IngrGeom3D.Line3d
    Set oRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(9), Port2S(10), Port2S(11), -BFLA(0), -BFLA(1), BFLA(2))
'   Construct branch rear right tangent line
    Dim oBRRTLine           As IngrGeom3D.Line3d
    Set oBRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(3), -BFLA(4), BFLA(5), Port3S(6), Port3S(7), Port3S(8))
'   Construct Port3 Rear Edge
    Dim oP3REdge           As IngrGeom3D.Line3d
    Set oP3REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), Port3S(9), Port3S(10), Port3S(11))
'   Construct branch Rear left tangent line
    Dim oBRLTLine           As IngrGeom3D.Line3d
    Set oBRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(9), Port3S(10), Port3S(11), BFLA(3), -BFLA(4), BFLA(5))
'   Construct Rear left tangent line from the branch arc start point
    Dim oRLTLine           As IngrGeom3D.Line3d
    Set oRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(0), -BFLA(1), BFLA(2), Port1S(9), Port1S(10), Port1S(11))
'   Construct Port1 Rear Edge
    Dim oP1REdge           As IngrGeom3D.Line3d
    Set oP1REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(9), Port1S(10), Port1S(11), Port1S(6), Port1S(7), Port1S(8))

'   Construct entire Bottom header and branch curves
    Dim oRearBranchcurves           As Collection
    Dim objRearHBcurves         As IngrGeom3D.ComplexString3d
    Dim objRearHcurves As IngrGeom3D.ComplexString3d
    Set oRearBranchcurves = New Collection

'Checking for the Part Data Basis Property
    Set oTrayPart = oPartFclt
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    
    Select Case lCableTrayPartDataBasis

        Case Is <= 1, 63, 65
'           Construct branch front right curve
            Dim objBFRA  As IngrGeom3D.Arc3d
            Dim SBFRA   As AutoMath.DPosition
            Set SBFRA = New AutoMath.DPosition
            Dim EBFRA   As AutoMath.DPosition
            Set EBFRA = New AutoMath.DPosition
            Dim CBFRA   As AutoMath.DPosition
            Set CBFRA = New AutoMath.DPosition
            SBFRA.Set -BFLA(0), BFLA(1), BFLA(2)
            EBFRA.Set -BFLA(3), BFLA(4), BFLA(5)
            CBFRA.Set -BFLA(6), BFLA(7), BFLA(8)
            Set objBFRA = PlaceTrArcByCenter(SBFRA, EBFRA, CBFRA)

'           Construct branch front left curve
            Dim objBFLA  As IngrGeom3D.Arc3d
            Dim SBFLA   As AutoMath.DPosition
            Set SBFLA = New AutoMath.DPosition
            Dim EBFLA   As AutoMath.DPosition
            Set EBFLA = New AutoMath.DPosition
            Dim CBFLA   As AutoMath.DPosition
            Set CBFLA = New AutoMath.DPosition
            SBFLA.Set BFLA(3), BFLA(4), BFLA(5)
            EBFLA.Set BFLA(0), BFLA(1), BFLA(2)
            CBFLA.Set BFLA(6), BFLA(7), BFLA(8)
            Set objBFLA = PlaceTrArcByCenter(SBFLA, EBFLA, CBFLA)

            oFrontBranchcurves.Add oHFLine
            oFrontBranchcurves.Add oP2FEdge
            oFrontBranchcurves.Add oFRTLine
            oFrontBranchcurves.Add objBFRA
            oFrontBranchcurves.Add oBFRTLine
            oFrontBranchcurves.Add oP3FEdge
            oFrontBranchcurves.Add oBFLTLine
            oFrontBranchcurves.Add objBFLA
            oFrontBranchcurves.Add oFLTLine
            oFrontBranchcurves.Add oP1FEdge
            
        '   Construct branch Rear right curve
            Dim objBRRA  As IngrGeom3D.Arc3d
            Dim SBRRA   As AutoMath.DPosition
            Set SBRRA = New AutoMath.DPosition
            Dim EBRRA   As AutoMath.DPosition
            Set EBRRA = New AutoMath.DPosition
            Dim CBRRA   As AutoMath.DPosition
            Set CBRRA = New AutoMath.DPosition
            SBRRA.Set -BFLA(0), -BFLA(1), BFLA(2)
            EBRRA.Set -BFLA(3), -BFLA(4), BFLA(5)
            CBRRA.Set -BFLA(6), -BFLA(7), BFLA(8)
            Set objBRRA = PlaceTrArcByCenter(SBRRA, EBRRA, CBRRA)
            
        '   Construct branch Rear left curve Edge
            Dim objBRLA  As IngrGeom3D.Arc3d
            Dim SBRLA   As AutoMath.DPosition
            Set SBRLA = New AutoMath.DPosition
            Dim EBRLA   As AutoMath.DPosition
            Set EBRLA = New AutoMath.DPosition
            Dim CBRLA   As AutoMath.DPosition
            Set CBRLA = New AutoMath.DPosition
            SBRLA.Set BFLA(3), -BFLA(4), BFLA(5)
            EBRLA.Set BFLA(0), -BFLA(1), BFLA(2)
            CBRLA.Set BFLA(6), -BFLA(7), BFLA(8)
            Set objBRLA = PlaceTrArcByCenter(SBRLA, EBRLA, CBRLA)

            oRearBranchcurves.Add oHRLine
            oRearBranchcurves.Add oP2REdge
            oRearBranchcurves.Add oRRTLine
            oRearBranchcurves.Add objBRRA
            oRearBranchcurves.Add oBRRTLine
            oRearBranchcurves.Add oP3REdge
            oRearBranchcurves.Add oBRLTLine
            oRearBranchcurves.Add objBRLA
            oRearBranchcurves.Add oRLTLine
            oRearBranchcurves.Add oP1REdge
        
        Case 67, 69
        '   Construct branch front right line
            Dim oBFRL           As IngrGeom3D.Line3d
            Set oBFRL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(0), BFLA(1), BFLA(2), -BFLA(3), BFLA(4), BFLA(5))

        '   Construct branch front left line
            Dim oBFLL           As IngrGeom3D.Line3d
            Set oBFLL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(3), BFLA(4), BFLA(5), BFLA(0), BFLA(1), BFLA(2))

            oFrontBranchcurves.Add oHFLine
            oFrontBranchcurves.Add oP2FEdge
            oFrontBranchcurves.Add oFRTLine
            oFrontBranchcurves.Add oBFRL
            oFrontBranchcurves.Add oBFRTLine
            oFrontBranchcurves.Add oP3FEdge
            oFrontBranchcurves.Add oBFLTLine
            oFrontBranchcurves.Add oBFLL
            oFrontBranchcurves.Add oFLTLine
            oFrontBranchcurves.Add oP1FEdge
            
        '   Construct branch Rear right line
            Dim oBRRL           As IngrGeom3D.Line3d
            Set oBRRL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(0), -BFLA(1), BFLA(2), -BFLA(3), -BFLA(4), BFLA(5))
            
        '   Construct branch Rear left line
            Dim oBRLL           As IngrGeom3D.Line3d
            Set oBRLL = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(3), -BFLA(4), BFLA(5), BFLA(0), -BFLA(1), BFLA(2))
            
            oRearBranchcurves.Add oHRLine
            oRearBranchcurves.Add oP2REdge
            oRearBranchcurves.Add oRRTLine
            oRearBranchcurves.Add oBRRL
            oRearBranchcurves.Add oBRRTLine
            oRearBranchcurves.Add oP3REdge
            oRearBranchcurves.Add oBRLTLine
            oRearBranchcurves.Add oBRLL
            oRearBranchcurves.Add oRLTLine
            oRearBranchcurves.Add oP1REdge

        Case Else
            GoTo ErrorLabel:

    End Select

'   Construct front-tee
    Dim StartBC As AutoMath.DPosition
    Set StartBC = New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set objHBcurves = PlaceTrCString(StartBC, oFrontBranchcurves)

    Dim ObjFrontTeeSection As IngrGeom3D.Plane3d
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 1, 0
    Set ObjFrontTeeSection = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjFrontTeeSection.AddBoundary(objHBcurves)
'   Remove cable tray front-tee Header and Branch lines
    Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontTeeSection
    Set ObjFrontTeeSection = Nothing
    Set oDirProj = Nothing
    Set StartBC = Nothing

'   Construct rear-tee
    Dim StartBCRear   As AutoMath.DPosition
    Set StartBCRear = New AutoMath.DPosition
    StartBCRear.Set Port1S(6), Port1S(7), Port1S(8)
    Set objRearHBcurves = PlaceTrCString(StartBCRear, oRearBranchcurves)
    
    Dim ObjRearTeeSection As IngrGeom3D.Plane3d
    Dim oDirProjRear        As AutoMath.DVector
    Set oDirProjRear = New AutoMath.DVector
    oDirProjRear.Set 0, 1, 0
    Set ObjRearTeeSection = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(6), Port1S(7), Port1S(8), oDirProjRear.x, oDirProjRear.y, oDirProjRear.z)
    Call ObjRearTeeSection.AddBoundary(objRearHBcurves)
'   Remove cable tray bottom Header and Branch lines
    Set ObjtopHBcurves = objRearHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
    Set StartBCRear = Nothing

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearTeeSection
    Set ObjRearTeeSection = Nothing
    Set oDirProjRear = Nothing

'   Remove branch curves
    Set oBFLTLine = Nothing
    Set objBFLA = Nothing
    Set objBFRA = Nothing
    Set objBRLA = Nothing
    Set objBRRA = Nothing
    Set oBRLTLine = Nothing
    Set oBRRTLine = Nothing
    Set oBFRTLine = Nothing
    Set oFLTLine = Nothing
    Set oFRTLine = Nothing
    Set oHFLine = Nothing
    Set oHRLine = Nothing
    Set oP1FEdge = Nothing
    Set oP1REdge = Nothing
    Set oP2FEdge = Nothing
    Set oP2REdge = Nothing
    Set oP3FEdge = Nothing
    Set oP3REdge = Nothing
    Set oRLTLine = Nothing
    Set oRRTLine = Nothing
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

'   Set the output
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

'   Set the output
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
    oDir.Set 0, 0, 1
    oRadialOrient.Set -1, 0, 0
    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort3 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oGeomFactory = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

