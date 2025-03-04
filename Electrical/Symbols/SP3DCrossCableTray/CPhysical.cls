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
'   Creation Date:  Thursday, Feb 28 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is prepared in the following way:
' ObjTrayBottom - Using the boundary curves
' ObjTopBranchPort1Side, ObjTopBranchPort2Side, ObjBottomBranchPort1Side, ObjBottomBranchPort2Side from
' Projection of a complex string (made of Tangent + Arc + Tangent)
' Three objCableTrayPort outputs by using 'CreateCableTrayPort' function
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----                ------------------
'  09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'  06.May.2004     MS                 TR 58960 Removed hardcoded TangentLength value.
'                                     Now the TangentLength is retrieved from oTrayPart
'  23.Nov.2004     MS                 CR 66680 Added input validation for TangentLength property for Tee and Cross
'                                     symbols (SP3DTeeCableTray, SP3DVTeeUpCableTray, SP3DVTeeDownCableTray, SP3DCrossCableTray)
'  08.SEP.2006     KKC                DI-95670 Replace names with initials in all revision history sheets and symbols
'  11.June.2007    KKC                DM-121367  Implemented part data basis for the Square Outside Corners.
'  14.JUN.2007     dkl                CR-114880 Implemented part data basis for segmented Cross. (Refer Horizontal Four Way, Zhen Hua Catalog.)
'  3.JUL.2007      dkl                CR-122171 Implemented part data bases 113 for Schneider Electric Catalog.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const MIN_VALUE = 0.0001

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
    Dim lCableTrayPartDataBasis As Integer
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parCTDimension = arrayOfInputs(2)
'    parHeaderLength = arrayOfInputs(3)
'    parTangentLength1 = arrayOfInputs(4)
'    parTangentLength2 = arrayOfInputs(5)
'    parTangentLength3 = arrayOfInputs(6)
'    parTangentLength4 = arrayOfInputs(7)
    
    iOutput = 0
    
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    Dim oPort3 As AutoMath.DPosition 'Port 3 center point
    Set oPort3 = New AutoMath.DPosition
    Dim oPort4 As AutoMath.DPosition 'Port 4 center point
    Set oPort4 = New AutoMath.DPosition

    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    Dim dTangentLength3 As Double   'Tangent length along Port 3.
    Dim dTangentLength4 As Double   'Tangent length along Port 4.

'   Retrieve Port 1 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HD As Double
    Dim HW As Double
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    
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
'   1) Curved Horizontal Four Way, Symmetrical, specified by Bend Radius and Tangent Length (Default implementation).
'   2) Bend Radius Dimension Exclusion basis.
'   3) Curved Horizontal Four Way, Asymmetrical, specified by Bend Radius Tangent Length 1, Tangent Length 2, Tangent Length 3 and Tangent Length 4.
'   4) Curved Horizontal Four Way, Symmetrical, specified by Header Length, Bend Radius and Tangent Length.
'   5) Curved Horizontal Four Way, Asymmetrical, specified by Header Length, Bend Radius, Tangent Length 1, Tangent Length 2, Tangent Length 3 and Tangent Length 4.
'   6) Segmented Horizontal Four Way, Symmetrical, specified by Header Length, Bend Radius and Tangent Length.
'   7) Segmented Horizontal Four Way, Asymmetrical, specified by Header Length, Bend Radius, Tangent Length 1, Tangent Length 2, Tangent Length 3 and Tangent Length 4.
'   8) Curved Horizontal Four Way, specified by Header Length and Bend Radius.

    Select Case lCableTrayPartDataBasis
        Case Is <= 1, 19     'Curved Horizontal Four Way, Symmetrical, specified by
                             'Bend Radius and Tangent Length (Default implementation).
            parTangentLength = oTrayPart.TangentLength
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            dTangentLength4 = parTangentLength
                        
            oPort1.Set -(dTangentLength1 + parBendRadius + HW), 0, 0
            oPort2.Set (dTangentLength2 + parBendRadius + HW), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            oPort4.Set 0, -(dTangentLength4 + parBendRadius + HW), 0
            
        Case 5      'Bend Radius Dimension Exclusion basis.
            Dim parCTDimension As Double
            parCTDimension = arrayOfInputs(2)
            parTangentLength = oTrayPart.TangentLength
            
            oPort1.Set -(parCTDimension + HW), 0, 0
            oPort2.Set (parCTDimension + HW), 0, 0
            oPort3.Set 0, (parCTDimension + HW), 0
            oPort4.Set 0, -(parCTDimension + HW), 0
            
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            dTangentLength4 = parTangentLength
            
        Case 21     'Curved Horizontal Four Way, Asymmetrical, specified by Bend Radius,
                    'Tangent Length 1, Tangent Length 2, Tangent Length 3 and Tangent Length 4.
            parBendRadius = oTrayPart.BendRadius
            
            dTangentLength1 = arrayOfInputs(4)
            dTangentLength2 = arrayOfInputs(5)
            dTangentLength3 = arrayOfInputs(6)
            dTangentLength4 = arrayOfInputs(7)
            
            oPort1.Set -(dTangentLength1 + parBendRadius + HW), 0, 0
            oPort2.Set (dTangentLength2 + parBendRadius + HW), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            oPort4.Set 0, -(dTangentLength4 + parBendRadius + HW), 0
            
        Case 23, 27     '23 - Curved Horizontal Four Way, Symmetrical, specified by
                        ' Header Length, Bend Radius and Tangent Length.
                        '27 - Segmented Horizontal Four Way, Symmetrical, specified by
                        'Header Length, Bend Radius and Tangent Length.
            Dim parHeaderLength As Double
            parHeaderLength = arrayOfInputs(3)
            parTangentLength = oTrayPart.TangentLength
            parBendRadius = oTrayPart.BendRadius
            
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            dTangentLength4 = parTangentLength
            
            oPort1.Set -(dTangentLength1 + parBendRadius + HW), 0, 0
            oPort2.Set parHeaderLength - (dTangentLength1 + parBendRadius + HW), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            oPort4.Set 0, -parHeaderLength + (dTangentLength3 + parBendRadius + HW), 0
            
        Case 25, 29     '25 - Curved Horizontal Four Way, Asymmetrical, specified by Header Length,
                        'Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
                        '29 - Segmented Horizontal Four Way, Asymmetrical, specified by Header Length,
                        ' Bend Radius, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
            parHeaderLength = arrayOfInputs(3)
            dTangentLength1 = arrayOfInputs(4)
            dTangentLength2 = arrayOfInputs(5)
            dTangentLength3 = arrayOfInputs(6)
            dTangentLength4 = arrayOfInputs(7)
            
            parBendRadius = oTrayPart.BendRadius
                        
            oPort1.Set -(dTangentLength1 + parBendRadius + HW), 0, 0
            oPort2.Set parHeaderLength - (dTangentLength1 + parBendRadius + HW), 0, 0
            oPort3.Set 0, (dTangentLength3 + parBendRadius + HW), 0
            oPort4.Set 0, -parHeaderLength + (dTangentLength3 + parBendRadius + HW), 0
        
        Case 113    'Curved Horizontal Four Way, specified by Header Length and Bend Radius.
            parHeaderLength = arrayOfInputs(3)
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = (parHeaderLength - parActualWidth - 2 * parBendRadius) / 2
            dTangentLength2 = dTangentLength1
            dTangentLength3 = dTangentLength1
            dTangentLength4 = dTangentLength1
                        
            oPort1.Set -(dTangentLength1 + parBendRadius + HW), 0, 0
            oPort2.Set dTangentLength2 + parBendRadius + HW, 0, 0
            oPort3.Set 0, dTangentLength3 + parBendRadius + HW, 0
            oPort4.Set 0, -(dTangentLength4 + parBendRadius + HW), 0
            
        Case Else
            GoTo ErrorLabel:
    
    End Select

    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = MIN_VALUE
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = MIN_VALUE
    If CmpDblLessThanOrEqualTo(dTangentLength3, 0) Then dTangentLength3 = MIN_VALUE
    If CmpDblLessThanOrEqualTo(dTangentLength4, 0) Then dTangentLength4 = MIN_VALUE

' Insert your code for output 1: Tray Bottom
    Dim Port1S(0 To 11)  As Double
    Dim Port2S(0 To 11)  As Double
    Dim Port3S(0 To 11)  As Double
    Dim Port4S(0 To 11)  As Double
    Dim BLTA(0 To 8)  As Double
    
'   Tray Port 1 U-shape points positions
'   Tray top edge
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
'   Tray top edge
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD
    
'   Port 2 position
        
'   Tray Port 2 U-shape points positions
'   Tray top edge
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
'   Tray top edge
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD
    
'   Port 3 position
     
'   Tray Port 3 U-shape points positions
'   Tray top edge
    Port3S(0) = oPort3.x - HW
    Port3S(1) = oPort3.y
    Port3S(2) = oPort3.z + HD
'   Tray bottom
    Port3S(3) = oPort3.x - HW
    Port3S(4) = oPort3.y
    Port3S(5) = oPort3.z - HD
    
    Port3S(6) = oPort3.x + HW
    Port3S(7) = oPort3.y
    Port3S(8) = oPort3.z - HD
'   Tray top edge
    Port3S(9) = oPort3.x + HW
    Port3S(10) = oPort3.y
    Port3S(11) = oPort3.z + HD
    
'   Port 4 position
    
'   Tray Port 4 U-shape points positions
'   Tray top edge
    Port4S(0) = oPort4.x - HW
    Port4S(1) = oPort4.y
    Port4S(2) = oPort4.z + HD
'   Tray bottom
    Port4S(3) = oPort4.x - HW
    Port4S(4) = oPort4.y
    Port4S(5) = oPort4.z - HD
    
    Port4S(6) = oPort4.x + HW
    Port4S(7) = oPort4.y
    Port4S(8) = oPort4.z - HD
'   Tray top edge
    Port4S(9) = oPort4.x + HW
    Port4S(10) = oPort4.y
    Port4S(11) = oPort4.z + HD
    
'   Branch  Left Top Arc BLTA
    BLTA(0) = Port1S(3) + dTangentLength1
    BLTA(1) = Port1S(4)
    BLTA(2) = Port1S(5)
'   Tray bottom
    BLTA(3) = Port3S(3)
    BLTA(4) = Port3S(4) - dTangentLength3
    BLTA(5) = Port3S(5)
    
    BLTA(6) = Port1S(3) + dTangentLength1
    BLTA(7) = Port3S(4) - dTangentLength3
    BLTA(8) = Port1S(5)

'   Construct Tray bottom: Branch curves
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
'   Construct Port1 Left Top Tangent line
    Dim oLTopTLine           As IngrGeom3D.Line3d
    Set oLTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), BLTA(0), BLTA(1), BLTA(2))
'   Construct branch left top tangent line
    Dim oBLTopTLine           As IngrGeom3D.Line3d
    Set oBLTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLTA(3), BLTA(4), BLTA(5), Port3S(3), Port3S(4), Port3S(5))
'   Construct branch right top tangent line
    Dim oBRTopTLine           As IngrGeom3D.Line3d
    Set oBRTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), -BLTA(3), BLTA(4), BLTA(5))
'   Construct  Right Top Tangent line at Port2
    Dim oRTopTLine           As IngrGeom3D.Line3d
    Set oRTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BLTA(0), BLTA(1), BLTA(2), Port2S(3), Port2S(4), Port2S(5))
'   Construct Port1 bottom curve
    Dim oPort1botLine           As IngrGeom3D.Line3d
    Set oPort1botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), Port1S(3), Port1S(4), Port1S(5))
'   Construct Port2 bottom curve
    Dim oPort2botLine           As IngrGeom3D.Line3d
    Set oPort2botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), Port2S(6), Port2S(7), Port2S(8))
'   Construct Port3 bottom curve
    Dim oPort3botLine           As IngrGeom3D.Line3d
    Set oPort3botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(6), Port3S(7), Port3S(8))
'   Construct Port4 bottom curve
    Dim oPort4botLine           As IngrGeom3D.Line3d
    Set oPort4botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(6), Port4S(7), Port4S(8), Port4S(3), Port4S(4), Port4S(5))
    
'   Bottom side Curves and Tangents
'   Construct  Right Bottom Tangent line at Port2
    Dim oRBottomTLine           As IngrGeom3D.Line3d
    Set oRBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), -BLTA(0), -BLTA(1), BLTA(2))
'   Construct  Right Bottom Tangent line at Port4
    Dim oBRBottomTLine           As IngrGeom3D.Line3d
    Set oBRBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BLTA(3), -BLTA(4), BLTA(5), Port4S(6), Port4S(7), Port4S(8))
'   Construct branch left Bottom tangent line
    Dim oBLBottomTLine           As IngrGeom3D.Line3d
    Set oBLBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(3), Port4S(4), Port4S(5), BLTA(3), -BLTA(4), BLTA(5))
'   Construct Port1 Left Bottom Tangent line
    Dim oLBottomTLine           As IngrGeom3D.Line3d
    Set oLBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLTA(0), -BLTA(1), BLTA(2), Port1S(6), Port1S(7), Port1S(8))

'   Construct collection for creating the bottom surface.
    Dim obranchcurves           As Collection
    Set obranchcurves = New Collection
    obranchcurves.Add oLTopTLine

'   Construct collection for creating Top Branch Port 1 Side.
    Dim oBp1scurve           As Collection
    Set oBp1scurve = New Collection
    oBp1scurve.Add oLTopTLine
    
'   Construct collection for creating Top Branch Port 2 Side.
    Dim oBp2scurve           As Collection
    Set oBp2scurve = New Collection
    oBp2scurve.Add oBRTopTLine

'   Construct collection for creating Bottom Branch Port 2 Side.
    Dim oBp3scurve           As Collection
    Set oBp3scurve = New Collection
    oBp3scurve.Add oRBottomTLine
    
'   Construct collection for creating Bottom Branch Port 1 Side.
    Dim oBp4scurve           As Collection
    Set oBp4scurve = New Collection
    oBp4scurve.Add oBLBottomTLine
    
        Select Case lCableTrayPartDataBasis
        Case Is <= 1, 19, 21, 23, 25, 113
            ' Construct branch left top curve
            Dim objBLTA  As IngrGeom3D.Arc3d
            Dim SBLTA   As New AutoMath.DPosition
            Dim EBLTA   As New AutoMath.DPosition
            Dim CBLTA   As New AutoMath.DPosition
            SBLTA.Set BLTA(0), BLTA(1), BLTA(2)
            EBLTA.Set BLTA(3), BLTA(4), BLTA(5)
            CBLTA.Set BLTA(6), BLTA(7), BLTA(8)
            Set objBLTA = PlaceTrArcByCenter(SBLTA, EBLTA, CBLTA)
            
            ' Construct branch right top curve
            Dim objBRTA  As IngrGeom3D.Arc3d
            Dim SBRTA   As New AutoMath.DPosition
            Dim EBRTA   As New AutoMath.DPosition
            Dim CBRTA   As New AutoMath.DPosition
            SBRTA.Set -BLTA(3), BLTA(4), BLTA(5)
            EBRTA.Set -BLTA(0), BLTA(1), BLTA(2)
            CBRTA.Set -BLTA(6), BLTA(7), BLTA(8)
            Set objBRTA = PlaceTrArcByCenter(SBRTA, EBRTA, CBRTA)
        
            ' Construct branch left Bottom curve
            Dim objBLeftBottomA  As IngrGeom3D.Arc3d
            Dim SBLBottomTA   As New AutoMath.DPosition
            Dim EBLBottomTA   As New AutoMath.DPosition
            Dim CBLBottomTA   As New AutoMath.DPosition
            SBLBottomTA.Set BLTA(3), -BLTA(4), BLTA(5)
            EBLBottomTA.Set BLTA(0), -BLTA(1), BLTA(2)
            CBLBottomTA.Set BLTA(6), -BLTA(7), BLTA(8)
            Set objBLeftBottomA = PlaceTrArcByCenter(SBLBottomTA, EBLBottomTA, CBLBottomTA)
            
            ' Construct branch right bottom curve
            Dim objBRightBottomA  As IngrGeom3D.Arc3d
            Dim SBBottomTA   As New AutoMath.DPosition
            Dim EBBottomTA   As New AutoMath.DPosition
            Dim CBBottomTA   As New AutoMath.DPosition
            SBBottomTA.Set -BLTA(0), -BLTA(1), BLTA(2)
            EBBottomTA.Set -BLTA(3), -BLTA(4), BLTA(5)
            CBBottomTA.Set -BLTA(6), -BLTA(7), BLTA(8)
            Set objBRightBottomA = PlaceTrArcByCenter(SBBottomTA, EBBottomTA, CBBottomTA)
            
            'Adding curves for the Tray Bottom.
            obranchcurves.Add objBLTA
            obranchcurves.Add oBLTopTLine
            obranchcurves.Add oPort3botLine
            obranchcurves.Add oBRTopTLine
            obranchcurves.Add objBRTA
            obranchcurves.Add oRTopTLine
            obranchcurves.Add oPort2botLine
            obranchcurves.Add oRBottomTLine
            obranchcurves.Add objBRightBottomA
            obranchcurves.Add oBRBottomTLine
            obranchcurves.Add oPort4botLine
            obranchcurves.Add oBLBottomTLine
            obranchcurves.Add objBLeftBottomA
            
            'Adding curves for the Top Branch Port 1 Side.
            oBp1scurve.Add objBLTA
            
            'Adding curves for the Top Branch Port 2 Side.
            oBp2scurve.Add objBRTA
            
            'Adding curves for the Bottom Branch Port 2 Side.
            oBp3scurve.Add objBRightBottomA
            
            'Adding curves for the Bottom Branch Port 1 Side.
            oBp4scurve.Add objBLeftBottomA
            
            Set objBLTA = Nothing
            Set objBRTA = Nothing
            Set objBLeftBottomA = Nothing
            Set objBRightBottomA = Nothing
            
        Case 5, 27, 29
            '   Construct branch left inclined line
            Dim objLILT  As IngrGeom3D.Line3d
            Set objLILT = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                BLTA(0), BLTA(1), BLTA(2), BLTA(3), BLTA(4), BLTA(5))
                                
            ' Construct branch right inclined line
            Dim objRILT  As IngrGeom3D.Line3d
            Set objRILT = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                -BLTA(3), BLTA(4), BLTA(5), -BLTA(0), BLTA(1), BLTA(2))
                                
            '   Construct branch left Bottom inclined line
            Dim objLILB  As IngrGeom3D.Line3d
            Set objLILB = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                BLTA(3), -BLTA(4), BLTA(5), BLTA(0), -BLTA(1), BLTA(2))
                                
            '   Construct branch right bottom inclined line
            Dim objRILB  As IngrGeom3D.Line3d
            Set objRILB = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                -BLTA(0), -BLTA(1), BLTA(2), -BLTA(3), -BLTA(4), BLTA(5))
            
            'Adding curves for the Tray Bottom.
            obranchcurves.Add objLILT
            obranchcurves.Add oBLTopTLine
            obranchcurves.Add oPort3botLine
            obranchcurves.Add oBRTopTLine
            obranchcurves.Add objRILT
            obranchcurves.Add oRTopTLine
            obranchcurves.Add oPort2botLine
            obranchcurves.Add oRBottomTLine
            obranchcurves.Add objRILB
            obranchcurves.Add oBRBottomTLine
            obranchcurves.Add oPort4botLine
            obranchcurves.Add oBLBottomTLine
            obranchcurves.Add objLILB
            
            'Adding curves for the Top Branch Port 1 Side.
            oBp1scurve.Add objLILT
           
           'Adding curves for the Top Branch Port 2 Side.
           oBp2scurve.Add objRILT
           
           'Adding curves for the Bottom Branch Port 2 Side.
           oBp3scurve.Add objRILB
           
           'Adding curves for the Bottom Branch Port 1 Side.
           oBp4scurve.Add objLILB

            Set objLILT = Nothing
            Set objRILT = Nothing
            Set objLILB = Nothing
            Set objRILB = Nothing
            
        Case Else
            GoTo ErrorLabel:
    End Select
       
    'Adding curves for the Tray Bottom.
    obranchcurves.Add oLBottomTLine
    obranchcurves.Add oPort1botLine

    'Adding curves for the Top Branch Port 1 Side.
    oBp1scurve.Add oBLTopTLine

    'Adding curves for the Top Branch Port 2 Side.
    oBp2scurve.Add oRTopTLine
    
    'Adding curves for the Bottom Branch Port 2 Side.
    oBp3scurve.Add oBRBottomTLine
    
    'Adding curves for the Bottom Branch Port 1 Side.
    oBp4scurve.Add oLBottomTLine
    
'   Construct entire Bottom surface.
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    
    Dim ObjTrayBottom As IngrGeom3D.Plane3d
    Set ObjTrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                    Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjTrayBottom.AddBoundary(objHBcurves)
       
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrayBottom
    Set ObjTrayBottom = Nothing
    Set objHBcurves = Nothing
    Set obranchcurves = Nothing
    
' Insert your code for output 2: TopBranchPort1Side
    Dim objBp1scurve         As IngrGeom3D.ComplexString3d
    Set objBp1scurve = PlaceTrCString(StartBC, oBp1scurve)
    
    Dim ObjTopBranchPort1Side As IJDObject
    Set ObjTopBranchPort1Side = PlaceProjection(m_OutputColl, objBp1scurve, oDirProj, _
                                                                    parActualDepth, True)
        
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBranchPort1Side
    Set ObjTopBranchPort1Side = Nothing
    Set objBp1scurve = Nothing
    Set oBp1scurve = Nothing
    
' Insert your code for output 3: TopBranchPort2Side
    Dim StartBC2   As New AutoMath.DPosition
    StartBC2.Set Port3S(6), Port3S(7), Port3S(8)
    
    Dim objBp2scurve         As IngrGeom3D.ComplexString3d
    Set objBp2scurve = PlaceTrCString(StartBC2, oBp2scurve)
    
    Dim ObjTopBranchPort2Side As IJDObject
    Set ObjTopBranchPort2Side = PlaceProjection(m_OutputColl, objBp2scurve, oDirProj, _
                                                                    parActualDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBranchPort2Side
    Set ObjTopBranchPort2Side = Nothing
    Set objBp2scurve = Nothing
    Set oBp2scurve = Nothing
    
' Insert your code for output 4: BottomBranchPort2Side
    Dim StartBC3   As New AutoMath.DPosition
    StartBC3.Set Port2S(6), Port2S(7), Port2S(8)
    
    Dim objBp3scurve         As IngrGeom3D.ComplexString3d
    Set objBp3scurve = PlaceTrCString(StartBC3, oBp3scurve)
    
    Dim ObjBottomBranchPort2Side As IJDObject
    Set ObjBottomBranchPort2Side = PlaceProjection(m_OutputColl, objBp3scurve, oDirProj, _
                                                                    parActualDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomBranchPort2Side
    Set ObjBottomBranchPort2Side = Nothing
    Set objBp3scurve = Nothing
    Set oBp3scurve = Nothing
    
' Insert your code for output 5: BottomBranchPort1Side
    Dim StartBC4   As New AutoMath.DPosition
    StartBC4.Set Port4S(3), Port4S(4), Port4S(5)
    
    Dim objBp4scurve         As IngrGeom3D.ComplexString3d
    Set objBp4scurve = PlaceTrCString(StartBC4, oBp4scurve)
    
    Dim ObjBottomBranchPort1Side As IJDObject
    Set ObjBottomBranchPort1Side = PlaceProjection(m_OutputColl, objBp4scurve, oDirProj, parActualDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomBranchPort1Side
    Set ObjBottomBranchPort1Side = Nothing
    Set oDirProj = Nothing
    Set objBp4scurve = Nothing
    Set oBp4scurve = Nothing
    Set oLTopTLine = Nothing
    Set oBLTopTLine = Nothing
    Set oBRTopTLine = Nothing
    Set oRTopTLine = Nothing
    Set oPort2botLine = Nothing
    Set oPort1botLine = Nothing
    Set oRBottomTLine = Nothing
    Set oBRBottomTLine = Nothing
    Set oBLBottomTLine = Nothing
    Set oLBottomTLine = Nothing
    Set oPort3botLine = Nothing
    Set oPort4botLine = Nothing
    
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

' Place Port 4
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition

    oDir.Set 0, -1, 0
    oRadialOrient.Set 0, 0, 1

    oPortLocation.Set oPort4.x - dInsertionDepth * oDir.x, oPort4.y - dInsertionDepth * oDir.y, oPort4.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 4, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort4 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
