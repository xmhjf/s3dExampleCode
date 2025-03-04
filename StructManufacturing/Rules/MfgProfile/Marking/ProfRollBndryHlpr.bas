Attribute VB_Name = "ProfRollBndryHlpr"
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileMarking
' Module: Profile Roll Boundary Helper
'
'*******************************************************************************

Option Explicit

Private Const MODULE As String = "ProfRollBndryHlpr::"

' NB: These map to the output as expected in the XML schema.
Private Const REND_BEND_LEFT As String = "L"
Private Const REND_BEND_RIGHT As String = "R"
Private Const REND_BEND_INFLECTION As String = "I"
Private Const REND_BEND_BOTH As String = "B"
Private Const REND_BEND_CURVE As String = "C"

Public Enum eSplineIntervalMarkType
    DoNotMark = 0
    ChordLength = 1
    GirthLength = 2
    Girth_ExtendLastByTangent = 3
    Girth_SetLastRadiusToPrev = 4
    Girth_SweepBackLastByIntvl = 5
End Enum

Private Const M_PI As Double = 3.14159265358979
Private Const SpillOverSplinePortion As Double = 0.15 ' 150 mm

Private Function SKIP_ROLL_REGION(ByVal RollRadius As Double, ByVal SweepAngle As Double) As Boolean
    'Skip if Roll radius is more than 50 metres or sweep angle is less than 5 degrees (0.087 radians)
    SKIP_ROLL_REGION = (Abs(RollRadius) > 50 Or SweepAngle < 5 * (M_PI / 180))
End Function
 
Private Function InvCosine(x As Double) As Double
    If Abs(x) > 1# Then
        Exit Function
    ElseIf Abs(x) < 0.001 Then
        InvCosine = M_PI / 2#
    Else
        InvCosine = Atn(-x / Sqr(-x * x + 1)) + M_PI / 2#
    End If
End Function

Private Function CheckForNewIntvlPoints(PtColl As IJElements, _
                                        NumNew As Long, _
                                        NewPoints() As Double) As IJDPosition
    Const METHOD = "CheckForNewIntvlPoints"
    On Error GoTo ErrorHandler
    
    If NumNew < 1 Then Exit Function

    Dim lbix As Long, ubix As Long
    lbix = LBound(NewPoints)
    ubix = UBound(NewPoints)
    If ubix - lbix + 1 <> 3 * NumNew Then Exit Function

    Dim i As Long
    For i = 0 To NumNew - 1
        Dim IdxOffset As Long
        IdxOffset = i * 3 + lbix
        
        Dim NewPt As New DPosition
        NewPt.Set NewPoints(IdxOffset), NewPoints(IdxOffset + 1), NewPoints(IdxOffset + 2)
        
        Dim PtExists As Boolean
        PtExists = False
        
        Dim IthPt As IJDPosition
        For Each IthPt In PtColl
            If IthPt.DistPt(NewPt) < 0.000001 Then
                PtExists = True
                Exit For
            End If
        Next
        
        If PtExists Then
            Set NewPt = Nothing
        Else
            Set CheckForNewIntvlPoints = NewPt
            Exit For
        End If
    Next

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function

Private Function CurvePortionBetTwoDists(oCS As IJComplexString, _
                                         ByVal Dist1 As Double, _
                                         ByVal Dist2 As Double) As IJComplexString
                                         
    Const METHOD As String = "RollBoundaries:CurvePortionBetTwoDists"
    On Error GoTo ErrorHandler

    Dim oCurve As IJCurve
    Set oCurve = oCS
    
    Dim Sx As Double, Sy As Double, Sz As Double
    Dim Ex As Double, Ey As Double, Ez As Double
    oCurve.EndPoints Sx, Sy, Sz, Ex, Ey, Ez
    
    Dim CurveLen As Double
    CurveLen = oCurve.length
    
    Dist1 = Abs(Dist1)
    Dist2 = Abs(Dist2)
    
    Dim oStartPos As New DPosition
    oStartPos.Set Sx, Sy, Sz
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oTrimPt1 As IJDPosition
    If Dist1 < 0.001 Then
        Set oTrimPt1 = oStartPos
    ElseIf Abs(CurveLen - Dist1) < 0.001 Or Dist1 > CurveLen Then
        Set oTrimPt1 = New DPosition
        oTrimPt1.Set Ex, Ey, Ez
    Else
        Set oTrimPt1 = oMfgGeomHelper.GetPointAtDistAlongCurve(oCS, oStartPos, Dist1)
    End If
    
    Dim oTrimPt2 As IJDPosition
    If Dist2 < 0.001 Then
        Set oTrimPt2 = oStartPos
    ElseIf Abs(CurveLen - Dist2) < 0.001 Or Dist2 > CurveLen Then
        Set oTrimPt2 = New DPosition
        oTrimPt2.Set Ex, Ey, Ez
    Else
        Set oTrimPt2 = oMfgGeomHelper.GetPointAtDistAlongCurve(oCS, oStartPos, Dist2)
    End If
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    
    Dim oTrimmedCS As IJComplexString
    oMfgMGHelper.CloneComplexString oCS, oTrimmedCS
    
    oMfgMGHelper.TrimCurveByPoints oTrimmedCS, oTrimPt1, oTrimPt2
    
    Set CurvePortionBetTwoDists = oTrimmedCS
    
CleanUp:
    Set oMfgGeomHelper = Nothing
    Set oMfgMGHelper = Nothing
    Set oTrimPt1 = Nothing
    Set oTrimPt2 = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function

Private Sub GetBowStringDepthsBetweenIntervals(oBigCS As IJComplexString, _
                                               oStartPos As IJDPosition, _
                                               oEndPos As IJDPosition, _
                                               ChordHeight As Double, _
                                               ChordDistance As Double, _
                                               ApproxRollRadius As Double, _
                                               ApproxSweepAngle As Double)

    Const METHOD = "GetBowStringDepthsBetweenIntervals"
    On Error GoTo ErrorHandler

    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    
    Dim oTrimmedCS As IJComplexString
    oMfgMGHelper.CloneComplexString oBigCS, oTrimmedCS
    
    oMfgMGHelper.TrimCurveByPoints oTrimmedCS, oStartPos, oEndPos
    
#If 0 Then
    ' USE MIN BOUNDING BOX TO GET APPROX ARC PROPERTIES
    
    Dim UnitZ As New DVector
    UnitZ.Set 0#, 0#, 1#
    
    Dim ChordVec As IJDVector
    Set ChordVec = oEndPos.Subtract(oStartPos)
    ChordVec.length = 1
    
    Dim Yvec As IJDVector
    Set Yvec = UnitZ.Cross(ChordVec)
    
    Dim oGeomElem As IJElements
    Set oGeomElem = New JObjectCollection
    oGeomElem.Add oTrimmedCS
    
    Dim oThreeVecColl As IJElements
    Set oThreeVecColl = New JObjectCollection
    oThreeVecColl.Add ChordVec
    oThreeVecColl.Add Yvec
    oThreeVecColl.Add UnitZ
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim MinBoxPointColl As IJElements
    Set MinBoxPointColl = oMfgGeomHelper.GetGeometryMinBoxByVectors(oGeomElem, oThreeVecColl)
    
    Dim Pt2 As IJDPosition, Pt3 As IJDPosition
    Set Pt2 = MinBoxPointColl.Item(2)
    Set Pt3 = MinBoxPointColl.Item(3)
    
    ChordHeight = Pt2.DistPt(Pt3)
    ChordDistance = oStartPos.DistPt(oEndPos)
    ApproxRollRadius = (ChordHeight / 2) + ((ChordDistance ^ 2) / (8 * ChordHeight))
    ApproxSweepAngle = 2 * InvCosine(1# - ChordHeight / ApproxRollRadius)
    
    Set oMfgGeomHelper = Nothing
    Set UnitZ = Nothing
    Set ChordVec = Nothing
    Set Yvec = Nothing
    Set oGeomElem = Nothing
    Set oThreeVecColl = Nothing
    Set Pt2 = Nothing
    Set Pt3 = Nothing
    Set MinBoxPointColl = Nothing
#End If

    ' GET APPROX ARC FROM 3 POINTS OF CURVE PORTION
    
    Dim oTrimCrv As IJCurve
    Set oTrimCrv = oTrimmedCS
    
    Dim x As Double, y As Double, z As Double
    oTrimCrv.PositionFRatio 0.5, x, y, z
    
    Dim oMidPos As New DPosition
    oMidPos.Set x, y, z
    
    If (2 * oMidPos.x - oStartPos.x - oEndPos.x) ^ 2 + _
       (2 * oMidPos.y - oStartPos.y - oEndPos.y) ^ 2 + _
       (2 * oMidPos.z - oStartPos.z - oEndPos.z) ^ 2 > 0.000000000004 _
    Then
        Dim oApproxArc As New Arc3d
        oApproxArc.DefineBy3Points oStartPos.x, oStartPos.y, oStartPos.z, _
                                   oMidPos.x, oMidPos.y, oMidPos.z, _
                                   oEndPos.x, oEndPos.y, oEndPos.z
        
        ApproxRollRadius = oApproxArc.Radius
        ApproxSweepAngle = oApproxArc.SweepAngle
        ChordDistance = oStartPos.DistPt(oEndPos)
        ChordHeight = ApproxRollRadius - Sqr((ApproxRollRadius ^ 2) - (ChordDistance ^ 2) / 4)
    Else
        ChordHeight = 0#
        ChordDistance = oStartPos.DistPt(oEndPos)
        ApproxSweepAngle = 0.000001 * ChordDistance / SpillOverSplinePortion
        ApproxRollRadius = ChordDistance / ApproxSweepAngle
    End If
    
    Set oTrimCrv = Nothing
    Set oMidPos = Nothing
    Set oApproxArc = Nothing
    
CleanUp:
    Set oMfgMGHelper = Nothing
    Set oTrimmedCS = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Sub

Private Function PointsOnCurveAtGirthLengthInterval(oCurve As IJCurve, oStartPos As IJDPosition, oEndPos As IJDPosition, DistIntvl As Double) As IJElements
    Const METHOD = "PointsOnCurveAtGirthLengthInterval"
    On Error GoTo ErrorHandler

    Dim IntvlPoints As IJElements
    Set IntvlPoints = New JObjectCollection
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim CurDist As Double
    CurDist = DistIntvl
    
    Dim CurLen As Double
    CurLen = oCurve.length
    
    While CurDist < CurLen
        IntvlPoints.Add oMfgGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, CurDist)
        CurDist = CurDist + DistIntvl
    Wend
    
    Set PointsOnCurveAtGirthLengthInterval = IntvlPoints
    
CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function

Private Function PointsOnCurveAtChordLengthInterval(oCurve As IJCurve, oStartPos As IJDPosition, oEndPos As IJDPosition, DistIntvl As Double) As IJElements
    Const METHOD = "PointsOnCurveAtChordLengthInterval"
    On Error GoTo ErrorHandler

    Dim IntvlPoints As IJElements
    Set IntvlPoints = New JObjectCollection
    
    IntvlPoints.Add oStartPos ' Add but remove later
    
    Dim Cx As Double, Cy As Double, Cz As Double
    Cx = oStartPos.x
    Cy = oStartPos.y
    Cz = oStartPos.z
    
    Do
        Dim IntxCircle As New Circle3d
        IntxCircle.DefineByCenterNormalRadius Cx, Cy, Cz, 0#, 0#, 1#, DistIntvl
    
        Dim NumIntx As Long, NumOvlps As Long
        Dim IntxPoints() As Double
        Dim IntxResult As Geom3dIntersectConstants
        oCurve.Intersect IntxCircle, NumIntx, IntxPoints, NumOvlps, IntxResult
        
        Dim NewCenter As IJDPosition
        Set NewCenter = CheckForNewIntvlPoints(IntvlPoints, NumIntx, IntxPoints)
        
        If Not NewCenter Is Nothing Then
            IntvlPoints.Add NewCenter
            NewCenter.Get Cx, Cy, Cz
        Else
            Exit Do
        End If
        
        Set IntxCircle = Nothing
        Set NewCenter = Nothing
    Loop While True

    ' Remove Start point from collection
    Dim IthPt As IJDPosition
    For Each IthPt In IntvlPoints
        If IthPt.DistPt(oStartPos) < 0.000001 Or IthPt.DistPt(oEndPos) < 0.000001 Then
            IntvlPoints.Remove IthPt
        End If
    Next
    
    Set PointsOnCurveAtChordLengthInterval = IntvlPoints
    
CleanUp:
    Set IntxCircle = Nothing
    Set IthPt = Nothing
    Set NewCenter = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function

Private Sub GetPointsAtDistIntervals(oCurve As IJCurve, _
                                     ByVal Dist1 As Double, _
                                     ByVal Dist2 As Double, _
                                     ByVal IntvlOption As eSplineIntervalMarkType, _
                                     ByVal DistIntvl As Double, _
                                     ChordHeights() As Double, _
                                     ChordLengths() As Double, _
                                     MidDistInvl() As Double, _
                                     FullInvlDists() As Double, _
                                     IntvlPoints As IJElements, _
                                     RegionRollRadius() As Double, _
                                     RegionSweepAngle() As Double)
                                     
    Const METHOD As String = "GetPointsAtDistIntervals"
    On Error GoTo ErrorHandler
    
    If Abs(Dist2) - Abs(Dist1) < DistIntvl Then Exit Sub

    Dim ReverseDists As Integer
    If Dist1 >= 0# And Dist2 >= 0# Then
        ReverseDists = 1
    ElseIf Dist1 <= 0# And Dist2 <= 0# Then
        ReverseDists = -1
    Else
        Dim MyErrString As String
        MyErrString = "Potential error! Check Roll boundary marks!" & vbCrLf & _
                      "Sought Curve portion between " & Dist1 & " and " & Dist2
        StrMfgLogError Err, MODULE, METHOD, MyErrString, "SMCustomWarningMessages", 2014, , "RULES"
        Exit Sub
    End If

    If oCurve.length < DistIntvl Then Exit Sub
    
    Dim Sx As Double, Sy As Double, Sz As Double
    Dim Ex As Double, Ey As Double, Ez As Double
    oCurve.EndPoints Sx, Sy, Sz, Ex, Ey, Ez
    
    Dim oStartPos As New DPosition
    oStartPos.Set Sx, Sy, Sz
    
    Dim oEndPos As New DPosition
    oEndPos.Set Ex, Ey, Ez
    
    If IntvlOption = eSplineIntervalMarkType.ChordLength Then
        Set IntvlPoints = PointsOnCurveAtChordLengthInterval(oCurve, oStartPos, oEndPos, DistIntvl)
    ElseIf IntvlOption = eSplineIntervalMarkType.GirthLength Or _
           IntvlOption = Girth_ExtendLastByTangent Or _
           IntvlOption = Girth_SetLastRadiusToPrev Or _
           IntvlOption = Girth_SweepBackLastByIntvl _
    Then
        Set IntvlPoints = PointsOnCurveAtGirthLengthInterval(oCurve, oStartPos, oEndPos, DistIntvl)
    End If
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim TreatLastPortionDifferent As Boolean
    TreatLastPortionDifferent = True
    
    If SpillOverSplinePortion > 0# And IntvlPoints.Count > 0 Then
        Dim ExtraSplineLength As Double
        oMfgGeomHelper.GetLengthBet2Points oCurve, _
                                           IntvlPoints.Item(IntvlPoints.Count), _
                                           oEndPos, ExtraSplineLength
                                           
        If ExtraSplineLength < SpillOverSplinePortion Then
            IntvlPoints.Remove IntvlPoints.Count
            TreatLastPortionDifferent = False
        End If
    End If

    ReDim ChordHeights(1 To IntvlPoints.Count + 1) As Double
    ReDim ChordLengths(1 To IntvlPoints.Count + 1) As Double
    ReDim MidDistInvl(1 To IntvlPoints.Count + 1) As Double
    ReDim RegionRollRadius(1 To IntvlPoints.Count + 1) As Double
    ReDim RegionSweepAngle(1 To IntvlPoints.Count + 1) As Double
    
    Dim oCS As New ComplexString3d
    oCS.AddCurve oCurve, False
    
    If IntvlPoints.Count = 0 Then
        GetBowStringDepthsBetweenIntervals oCS, _
                                           oStartPos, _
                                           oEndPos, _
                                           ChordHeights(1), _
                                           ChordLengths(1), _
                                           RegionRollRadius(1), _
                                           RegionSweepAngle(1)

        MidDistInvl(1) = (Dist1 + Dist2) * 0.5
    ElseIf IntvlPoints.Count > 0 Then
        ReDim FullInvlDists(1 To IntvlPoints.Count) As Double
        
        GetBowStringDepthsBetweenIntervals oCS, _
                                           oStartPos, _
                                           IntvlPoints.Item(1), _
                                           ChordHeights(1), _
                                           ChordLengths(1), _
                                           RegionRollRadius(1), _
                                           RegionSweepAngle(1)
                                                                  
        oMfgGeomHelper.GetLengthBet2Points oCurve, oStartPos, _
                                           IntvlPoints.Item(1), _
                                           FullInvlDists(1)
                                           
        MidDistInvl(1) = Dist1 + ReverseDists * FullInvlDists(1) * 0.5
        FullInvlDists(1) = Dist1 + ReverseDists * FullInvlDists(1)
        
        Dim i As Long
        For i = 1 To IntvlPoints.Count
            If i = IntvlPoints.Count Then
                Dim oMfgMGHelper As New MfgMGHelper
                Dim oLastSP As IJDPosition, oLastEP As IJDPosition

                Dim oModifiedLastCS As IJComplexString
                Dim LastCrvLen As Double
                Dim oTempCrv As IJCurve

                If Not TreatLastPortionDifferent Then
                    Set oModifiedLastCS = Nothing
                    Set oLastSP = Nothing
                    Set oLastEP = Nothing
                ElseIf IntvlOption = Girth_ExtendLastByTangent Then
                    oMfgMGHelper.CloneComplexString oCS, oModifiedLastCS
                    oMfgMGHelper.TrimCurveByPoints oModifiedLastCS, IntvlPoints.Item(i), oEndPos

                    Set oTempCrv = oModifiedLastCS
                    LastCrvLen = oTempCrv.length
                    Set oTempCrv = Nothing

                    If DistIntvl - LastCrvLen > 0.001 Then
                        Const ExtendByDistance As Integer = 1
                        Const ExtendTangentially As Integer = 0
                        Const ExtendTheEndSide As Integer = 1

                        Dim oExtendCS As IJComplexString
                        Set oExtendCS = oMfgGeomHelper.ExtraPolateCurve(oModifiedLastCS, _
                                                                        ExtendByDistance, _
                                                                        ExtendTangentially, _
                                                                        ExtendTheEndSide, _
                                                                        Nothing, Nothing, _
                                                                        DistIntvl - LastCrvLen)

                        Set oTempCrv = oExtendCS
                        oTempCrv.EndPoints Sx, Sy, Sz, Ex, Ey, Ez
                        Set oTempCrv = Nothing

                        Set oModifiedLastCS = oExtendCS
                        Set oExtendCS = Nothing

                        Set oLastSP = New DPosition
                        oLastSP.Set Sx, Sy, Sz

                        Set oLastEP = New DPosition
                        oLastEP.Set Ex, Ey, Ez
                    End If
                ElseIf IntvlOption = Girth_SetLastRadiusToPrev Then
                    oMfgGeomHelper.GetLengthBet2Points oCS, IntvlPoints.Item(i), oEndPos, LastCrvLen
                    
                    ' Set Radius of previous region to this one. Then calculate other properties
                    RegionRollRadius(i + 1) = RegionRollRadius(i)
                    RegionSweepAngle(i + 1) = LastCrvLen / RegionRollRadius(i + 1)
                    ChordLengths(i + 1) = oEndPos.DistPt(IntvlPoints.Item(i))
                    ChordHeights(i + 1) = RegionRollRadius(i + 1) - _
                                         Sqr((RegionRollRadius(i + 1) ^ 2) - (ChordLengths(i + 1) ^ 2) / 4)
                ElseIf IntvlOption = Girth_SweepBackLastByIntvl Then
                    oMfgGeomHelper.GetLengthBet2Points oCS, IntvlPoints.Item(i), oEndPos, LastCrvLen
                    
                    Set oLastSP = oMfgGeomHelper.GetPointAtDistAlongCurve(oCS, oEndPos, DistIntvl)
                    Set oLastEP = oEndPos.Clone

                    oMfgMGHelper.CloneComplexString oCS, oModifiedLastCS
                    oMfgMGHelper.TrimCurveByPoints oModifiedLastCS, oLastSP, oLastEP
                End If

                Set oMfgMGHelper = Nothing

                If Not (oModifiedLastCS Is Nothing Or oLastSP Is Nothing Or oLastEP Is Nothing) Then
                    GetBowStringDepthsBetweenIntervals oModifiedLastCS, _
                                                       oLastSP, oLastEP, _
                                                       ChordHeights(i + 1), _
                                                       ChordLengths(i + 1), _
                                                       RegionRollRadius(i + 1), _
                                                       RegionSweepAngle(i + 1)

                    ' Need to set these based on the actual residual portion, not the modified last portion
                    RegionSweepAngle(i + 1) = LastCrvLen / RegionRollRadius(i + 1)
                    ChordLengths(i + 1) = oEndPos.DistPt(IntvlPoints.Item(i))
                    ChordHeights(i + 1) = RegionRollRadius(i + 1) - _
                                          Sqr((RegionRollRadius(i + 1) ^ 2) - (ChordLengths(i + 1) ^ 2) / 4)

                    Set oModifiedLastCS = Nothing
                ElseIf Not TreatLastPortionDifferent Or IntvlOption <> Girth_SetLastRadiusToPrev Then
                    GetBowStringDepthsBetweenIntervals oCS, _
                                                   IntvlPoints.Item(i), _
                                                   oEndPos, _
                                                   ChordHeights(i + 1), _
                                                   ChordLengths(i + 1), _
                                                   RegionRollRadius(i + 1), _
                                                   RegionSweepAngle(i + 1)
                End If
                MidDistInvl(i + 1) = (FullInvlDists(i) + Dist2) / 2#
            Else
                GetBowStringDepthsBetweenIntervals oCS, _
                                                   IntvlPoints.Item(i), _
                                                   IntvlPoints.Item(i + 1), _
                                                   ChordHeights(i + 1), _
                                                   ChordLengths(i + 1), _
                                                   RegionRollRadius(i + 1), _
                                                   RegionSweepAngle(i + 1)
                                                                         
                oMfgGeomHelper.GetLengthBet2Points oCurve, oStartPos, _
                                                   IntvlPoints.Item(i + 1), _
                                                   FullInvlDists(i + 1)
                                           
                FullInvlDists(i + 1) = Dist1 + ReverseDists * FullInvlDists(i + 1)
                MidDistInvl(i + 1) = (FullInvlDists(i) + FullInvlDists(i + 1)) * 0.5
            End If
        Next
    End If
    
CleanUp:
    Set oModifiedLastCS = Nothing
    Set oCS = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Sub

Private Function IsRollPosBackwards(oProfRollPos As Collection) As Boolean
    Const METHOD = "IsRollPosBackwards"
    On Error GoTo ErrorHandler

    IsRollPosBackwards = False
    
    If oProfRollPos Is Nothing Then Exit Function
    
    Dim LastIdx As Long
    LastIdx = oProfRollPos.Count
    
    If LastIdx = 0 Then Exit Function
    
    Dim oFirst As IJDPosition
    Set oFirst = oProfRollPos.Item(1)
    
    Dim oLast As IJDPosition
    Set oLast = oProfRollPos.Item(LastIdx)
    
    IsRollPosBackwards = (oLast.x < oFirst.x)
    
CleanUp:
    Set oLast = Nothing
    Set oFirst = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function

Private Sub CheckBendRollCurveOrientation(oCurve As IJCurve, CurveLen As Double, AccumLen As Double, Mult As Integer)
    Const METHOD = "CheckBendRollCurveOrientation"
    On Error GoTo ErrorHandler

    If oCurve Is Nothing Then Exit Sub
    
    CurveLen = oCurve.length
    
    Dim Sx As Double, Sy As Double, Sz As Double, _
        Ex As Double, Ey As Double, Ez As Double
    oCurve.EndPoints Sx, Sy, Sz, Ex, Ey, Ez
    
    If Ex > 0 Then
        If Sx > Ex Then
            Mult = -1
            AccumLen = CurveLen
        Else
            Mult = 1
            AccumLen = 0
        End If
    Else
        If Sx > Ex Then
            Mult = -1
            AccumLen = 0
        Else
            Mult = 1
            AccumLen = CurveLen
        End If
    End If
    
CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Sub

Private Function IsArcInwardOrOutward(oCurve As IJCurve) As Double
    Const METHOD = "IsArcInwardOrOutward"
    On Error GoTo ErrorHandler
    
    'Find if Arc is going "inward" or "outward"
    Dim Sx As Double, Sy As Double, Sz As Double
    Dim Ex As Double, Ey As Double, Ez As Double
    oCurve.EndPoints Sx, Sy, Sz, Ex, Ey, Ez
    
    Dim Mx As Double, My As Double, Mz As Double
    oCurve.PositionFRatio 0.5, Mx, My, Mz
    
    Dim StartToEnd As New DVector
    StartToEnd.Set (Ex - Sx), (Ey - Sy), (Ez - Sz)
    
    Dim StartToMid As New DVector
    StartToMid.Set (Mx - Sx), (My - Sy), (Mz - Sz)
    
    Dim UnitNegZVec As New DVector
    UnitNegZVec.Set 0#, 0#, -1#
        
'    IsArcInwardOrOutward = StartToEnd.Angle(StartToMid, UnitZVec)
    IsArcInwardOrOutward = StartToEnd.Cross(StartToMid).Dot(UnitNegZVec)
   
CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function


Private Sub CreateRollBoundaryGeom2d(oResourceManager As Object, x As Double, dBotHeight As Double, dTopHeight As Double, _
                                     oGeomCol2dOut As IJMfgGeomCol2d, _
                                     UpSide As Long, RollBoundaryName As String, _
                                     SweepAngle As Double, RollRadius As Double, RollDirection As String, _
                                     Optional InflectionStr As String = "", _
                                     Optional subgeomtype As StrMfgGeometryType = STRMFG_UNDEFINED_GEOMTYPE)
    Const METHOD = "CreateRollBoundaryGeom2d"
    On Error GoTo ErrorHandler

    Dim oLine As IJLine
    Set oLine = New Line3d
    oLine.DefineBy2Points x, dBotHeight, 0, x, dTopHeight, 0

    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine, True
    
    Dim oNew2dGeom As IJMfgGeom2d
    Set oNew2dGeom = m_oGeom2dFactory.Create(oResourceManager)
    
    oNew2dGeom.PutGeometry oCS
    oNew2dGeom.PutGeometrytype STRMFG_ROLL_BOUNDARIES_MARK
    oNew2dGeom.FaceId = UpSide
    
    oGeomCol2dOut.AddGeometry 1, oNew2dGeom

    'Create a SystemMark object to store additional information
    Dim oSystemMark As IJMfgSystemMark
    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    
    'Set the marking side
    oSystemMark.SetMarkingSide UpSide
    oSystemMark.Set2dGeometry oNew2dGeom
    
    'QI for the MarkingInfo object on the SystemMark
    Dim oMarkingInfo As IJMarkingInfo
    Set oMarkingInfo = oSystemMark
        
    oMarkingInfo.name = RollBoundaryName
    oMarkingInfo.Radius = RollRadius
    oMarkingInfo.FittingAngle = SweepAngle
    oMarkingInfo.direction = RollDirection
    oMarkingInfo.FlangeDirection = InflectionStr
    
    If InflectionStr <> vbNullString And subgeomtype <> STRMFG_UNDEFINED_GEOMTYPE Then
        oNew2dGeom.PutSubGeometryType subgeomtype
    End If
    
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing

    Set oLine = Nothing
    Set oCS = Nothing
    Set oNew2dGeom = Nothing

CleanUp:
    Exit Sub

ErrorHandler:
End Sub

Private Function CreateRollBoundaryTextSupportMark(ByVal x As Double, ByVal y As Double, ByRef oResourceManager As IUnknown, ByVal GeomType As StrMfgGeometryType, ByVal UpSide As Long, ByRef oGeomCol2dOut As MfgGeomCol2d, ByVal RollRadius As Double, ByVal SweepAngle As Double, ByVal dWebThickness As Double) As IJMarkingInfo
    Dim oNew2dGeom As IJMfgGeom2d
    Dim oLine As IJLine
    Set oLine = New Line3d
    
    oLine.DefineBy2Points x - 0.005, y, 0, x + 0.005, y, 0
    
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine, True
    
    Set oNew2dGeom = m_oGeom2dFactory.Create(oResourceManager)
    oNew2dGeom.PutGeometry oCS
    oNew2dGeom.PutGeometrytype STRMFG_ROLL_BOUNDARIES_MARK
    'oNew2dGeom.PutSubGeometryType STRMFG_REF_MARK
    oNew2dGeom.FaceId = UpSide
    oNew2dGeom.IsSupportOnly = True
    
    oGeomCol2dOut.AddGeometry 1, oNew2dGeom
    
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMarkingInfo As IJMarkingInfo
    Dim oMoniker As IMoniker
    
    'Create a SystemMark object to store additional information
    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    Set oObjSystemMark = oSystemMark
                    
    'Set the marking side
    oSystemMark.SetMarkingSide SideA
    oSystemMark.Set2dGeometry oNew2dGeom
    
    'QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark
    
    oMarkingInfo.name = "RE"
    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "RollBoundary"
    oMarkingInfo.FittingAngle = SweepAngle
    oMarkingInfo.Radius = Abs(RollRadius)
    If RollRadius < 0 Then
'        oMarkingInfo.Radius = Abs(RollRadius) + (0.5 * dWebThickness)
        oMarkingInfo.direction = "down"
    Else
'        oMarkingInfo.Radius = Abs(RollRadius) - (0.5 * dWebThickness)
        oMarkingInfo.direction = "up"
    End If
    
    Set CreateRollBoundaryTextSupportMark = oMarkingInfo
    
    Set oSystemMark = Nothing
    Set oObjSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oMoniker = Nothing

    Set oLine = Nothing
    Set oCS = Nothing
    Set oNew2dGeom = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateProfileRollBoundaryMarks
' Purpose   : Return the set of roll boundaries that are the ouput of the unfold
'             If the profile is a knuckled, return no roll boundaries
'---------------------------------------------------------------------------------------

Public Function CreateProfileRollBoundaryMarks(Part As IJProfilePart, UpSide As Long, _
                                               MarkingInSplineRegion As eSplineIntervalMarkType, _
                                               SplineMarkInterval As Double, _
                                               Optional ReportBendDirAsRollerSide As Boolean = False) As IJMfgGeomCol2d
    Const METHOD = "CreateProfileRollBoundaryMarks"
    On Error GoTo ErrorHandler

    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part

    Dim oMfgPart As IJMfgProfilePart
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    ' Try to generate roll boundaries for knuckled profiles (hence comment below check)
'    If ((oProfileWrapper.CurvatureType = PROFILE_CURVATURE_BendKnuckleAlongFlange) Or _
'        (oProfileWrapper.CurvatureType = PROFILE_CURVATURE_BendKnuckleAlongWeb) Or _
'        (oProfileWrapper.CurvatureType = PROFILE_CURVATURE_KnuckledAlongFlange) Or _
'        (oProfileWrapper.CurvatureType = PROFILE_CURVATURE_KnuckledAlongWeb)) Then
'                Exit Function
'    End If
    
    Dim oMfgGeomCol2d As IJMfgGeomCol2d
    Set oMfgGeomCol2d = oMfgProfileWrapper.GetRollBoundaryLines
    
    If oMfgGeomCol2d Is Nothing Then
        'Since there are no roll boundaries we can leave the function
        GoTo CleanUp
    End If
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
      
    Dim oGeomCol2dOut As IJMfgGeomCol2d
    Set oGeomCol2dOut = m_oGeomCol2dFactory.Create(oResourceManager)

    Dim oMathGeom As IJDMfgGeomUtilWrapper
    Set oMathGeom = New GSCADMathGeom.MfgGeomUtilWrapper
    
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    
    Set oPartSupport.Part = oMfgPart.GetDetailedPart
    Set oProfilePartSupport = oPartSupport
    
    Dim dLowestPoint As Double, dHighestPoint As Double, dProfileHeight As Double, dFlangeWidth As Double
    oProfilePartSupport.GetThickness BottomFlange, dLowestPoint
    oProfilePartSupport.GetThickness TopFlange, dHighestPoint
    oProfilePartSupport.GetWebDepth dProfileHeight
    oProfilePartSupport.GetFlangeWidth dFlangeWidth
    
    Dim dWebThickness As Double
    oProfilePartSupport.GetThickness Web, dWebThickness
    
    Dim MarkTopHeight As Double
    MarkTopHeight = (dProfileHeight - dHighestPoint)
    
    Dim oGeom2d As IJMfgGeom2d
    Dim GeomType As StrMfgGeometryType
    
    Dim j As Integer
    For j = 1 To oMfgGeomCol2d.Getcount
        Set oGeom2d = oMfgGeomCol2d.GetGeometry(j)
        GeomType = oGeom2d.GetGeometryType
        If GeomType = STRMFG_PROFILE_ROLL_INFO Or _
           GeomType = STRMFG_PROFILE_BEND_INFO Then
            
            Dim oCS As IJComplexString
            Set oCS = oGeom2d.GetGeometry
            
            Dim Multiplier As Integer
            Dim WholeCurveLen As Double, AccumLen As Double
            CheckBendRollCurveOrientation oCS, WholeCurveLen, AccumLen, Multiplier
            
            Dim MarkBottomHeight As Double
            If GeomType = STRMFG_PROFILE_BEND_INFO Then
                MarkBottomHeight = dLowestPoint
            Else
                MarkBottomHeight = dLowestPoint + (dProfileHeight - dHighestPoint - dLowestPoint) / 2#
            End If
            
            Dim TextHeight As Double
            TextHeight = (MarkBottomHeight + MarkTopHeight) / 2#

            Dim i As Long
            For i = 1 To oCS.CurveCount
                Dim oCurve As IJCurve
                oCS.GetCurve i, oCurve
                
                Dim oPrevCurve As IJCurve, oPrevArc As IJArc
                If i > 1 Then
                    oCS.GetCurve i - 1, oPrevCurve
                    If TypeOf oPrevCurve Is IJLine Then
                        Set oPrevCurve = Nothing
                        Set oPrevArc = Nothing
                    ElseIf TypeOf oPrevCurve Is IJArc Then
                        Set oPrevArc = oPrevCurve
                    Else
                        ' Prev Curve is a spline, not an arc
                        Set oPrevArc = Nothing
                    End If
                Else
                    Set oPrevCurve = Nothing
                    Set oPrevArc = Nothing
                End If
                
                Dim oNextCurve As IJCurve, oNextArc As IJArc
                If i < oCS.CurveCount Then
                    oCS.GetCurve i + 1, oNextCurve
                    If TypeOf oNextCurve Is IJLine Then
                        Set oNextCurve = Nothing
                        Set oNextArc = Nothing
                    ElseIf TypeOf oNextCurve Is IJArc Then
                        Set oNextArc = oNextCurve
                    Else
                        ' Next curve is a spline, not an arc
                        Set oNextArc = Nothing
                    End If
                Else
                    Set oNextCurve = Nothing
                    Set oNextArc = Nothing
                End If
                
                Dim CurAccumLen As Double
                CurAccumLen = AccumLen + Multiplier * oCurve.length
                
                Dim RollDir As String
                Dim MarkOtherEndOfCurve As Boolean
                Dim CurveInOrOut As Integer
                
                Dim InflectStr As String
                InflectStr = vbNullString
                
                If TypeOf oCurve Is IJArc Then
                    ' PROCESS ARCS
                    
                    Dim oArc As IJArc
                    Set oArc = oCurve
                    
                    If SKIP_ROLL_REGION(oArc.Radius, oArc.SweepAngle) Then GoTo NextCurveInCS
                    
                    If IsArcInwardOrOutward(oCurve) > 0# Xor _
                       ReportBendDirAsRollerSide Then
                        CurveInOrOut = -1
                        RollDir = "down"
                    Else
                        CurveInOrOut = 1
                        RollDir = "up"
                    End If
                    
                    
                    If Not oPrevCurve Is Nothing Then
                        If IsArcInwardOrOutward(oCurve) * _
                           IsArcInwardOrOutward(oPrevCurve) < 0# Then
                            InflectStr = REND_BEND_INFLECTION
                        Else
                            InflectStr = REND_BEND_BOTH
                        End If
                    End If
                    
                    If InflectStr = vbNullString Then
                        If Multiplier < 0 Then
                            InflectStr = REND_BEND_LEFT
                        Else
                            InflectStr = REND_BEND_RIGHT
                        End If
                    End If
                    
                    CreateRollBoundaryGeom2d oResourceManager, AccumLen, MarkBottomHeight, MarkTopHeight, _
                                             oGeomCol2dOut, UpSide, "ROLLBOUNDARY-" & CStr(2 * i - 1), _
                                             oArc.SweepAngle, oArc.Radius, RollDir, InflectStr, STRMFG_UNDEFINED_GEOMTYPE
                    
                    CreateRollBoundaryTextSupportMark (AccumLen + CurAccumLen) / 2#, _
                                                      TextHeight, oResourceManager, _
                                                      GeomType, UpSide, oGeomCol2dOut, _
                                                      oArc.Radius * CurveInOrOut, _
                                                      oArc.SweepAngle, dWebThickness
                    
                    If oNextCurve Is Nothing Then
                        MarkOtherEndOfCurve = True
                    ElseIf Not oNextArc Is Nothing Then
                        If SKIP_ROLL_REGION(oNextArc.Radius, oNextArc.SweepAngle) Then
                            ' Next curve is an arc that will be skipped so mark this region's roll boundary
                            MarkOtherEndOfCurve = True
                        Else
                            MarkOtherEndOfCurve = False
                        End If
                    Else
                        ' Arc's RB shared with Spline => Arc gets preference
                        MarkOtherEndOfCurve = True
                    End If
                    
                    If MarkOtherEndOfCurve Then
                        If Multiplier < 0 Then
                            InflectStr = REND_BEND_RIGHT
                        Else
                            InflectStr = REND_BEND_LEFT
                        End If
                        
                        CreateRollBoundaryGeom2d oResourceManager, CurAccumLen, MarkBottomHeight, MarkTopHeight, _
                                                 oGeomCol2dOut, UpSide, "ROLLBOUNDARY-" & CStr(2 * i), _
                                                 oArc.SweepAngle, oArc.Radius, RollDir, InflectStr, STRMFG_UNDEFINED_GEOMTYPE
                    End If
                    
                ElseIf Not TypeOf oCurve Is IJLine And _
                       MarkingInSplineRegion <> DoNotMark And _
                       SplineMarkInterval > 0# Then
                       
                    ' Not curve, not line => B-Spline.  PROCESS B-SPLINES
                    
                    If IsArcInwardOrOutward(oCurve) > 0# Xor _
                       ReportBendDirAsRollerSide Then
                        CurveInOrOut = -1
                        RollDir = "down"
                    Else
                        CurveInOrOut = 1
                        RollDir = "up"
                    End If
                    
                    Dim BowStringDepths() As Double
                    Dim ChordLengths() As Double
                    Dim MidIntvlDists() As Double
                    Dim FullIntvlDists() As Double
                    Dim ApproxRadiusForInterval() As Double
                    Dim ApproxSweepForInterval() As Double
    
                    ReDim BowStringDepths(0) As Double
                    ReDim ChordLengths(0) As Double
                    ReDim MidIntvlDists(0) As Double
                    ReDim FullIntvlDists(0) As Double
                    ReDim ApproxRadiusForInterval(0) As Double
                    ReDim ApproxSweepForInterval(0) As Double
    
                    Dim IntervalPoints As IJElements
                    
                    GetPointsAtDistIntervals oCurve, _
                                             AccumLen, _
                                             CurAccumLen, _
                                             MarkingInSplineRegion, _
                                             SplineMarkInterval, _
                                             BowStringDepths, _
                                             ChordLengths, _
                                             MidIntvlDists, _
                                             FullIntvlDists, _
                                             IntervalPoints, _
                                             ApproxRadiusForInterval, _
                                             ApproxSweepForInterval
                                             
                    Dim Idx As Long
                    
                    If Not (LBound(BowStringDepths) = 0 And UBound(BowStringDepths) = 0) And _
                       (UBound(BowStringDepths) - LBound(BowStringDepths)) > -1 _
                    Then
                        For Idx = LBound(BowStringDepths) To UBound(BowStringDepths)
                            CreateRollBoundaryTextSupportMark MidIntvlDists(Idx), TextHeight, oResourceManager, _
                                                              GeomType, UpSide, oGeomCol2dOut, _
                                                              ApproxRadiusForInterval(Idx) * CurveInOrOut, _
                                                              ApproxSweepForInterval(Idx), dWebThickness

                        Next
                    End If
    
                    If Not oPrevCurve Is Nothing And oPrevArc Is Nothing Then
                        If IsArcInwardOrOutward(oCurve) * _
                           IsArcInwardOrOutward(oPrevCurve) < 0# Then
                            InflectStr = REND_BEND_INFLECTION
                        Else
                            InflectStr = REND_BEND_BOTH
                        End If
                    End If
                    
                    If InflectStr = vbNullString Then
                        If Multiplier < 0 Then
                            InflectStr = REND_BEND_LEFT
                        Else
                            InflectStr = REND_BEND_RIGHT
                        End If
                    End If
                    
                    Dim PrevArcNotMarked As Boolean
                    If oPrevArc Is Nothing Then
                        PrevArcNotMarked = True
                    ' VB does not have short circuit evaluation
                    ElseIf SKIP_ROLL_REGION(oPrevArc.Radius, oPrevArc.SweepAngle) Then
                        PrevArcNotMarked = True
                    Else
                        ' If previous curve was arc, then it would have marked its boundary
                        PrevArcNotMarked = False
                    End If

                    If PrevArcNotMarked Then
                        CreateRollBoundaryGeom2d oResourceManager, AccumLen, MarkBottomHeight, MarkTopHeight, _
                                                 oGeomCol2dOut, UpSide, "ROLLBOUNDARY-" & CStr(2 * i - 1), _
                                                 ApproxSweepForInterval(LBound(ApproxSweepForInterval)), _
                                                 ApproxRadiusForInterval(LBound(ApproxRadiusForInterval)), _
                                                 RollDir, InflectStr, STRMFG_UNDEFINED_GEOMTYPE
                    End If
                    
                    If Not (LBound(FullIntvlDists) = 0 And UBound(FullIntvlDists) = 0) And _
                       (UBound(FullIntvlDists) - LBound(FullIntvlDists)) > -1 _
                    Then
                        For Idx = LBound(FullIntvlDists) To UBound(FullIntvlDists)
                            CreateRollBoundaryGeom2d oResourceManager, FullIntvlDists(Idx), _
                                                     MarkBottomHeight, MarkTopHeight, oGeomCol2dOut, _
                                                     UpSide, "SPLINE-INTERVAL-MARK-" & i & "-" & Idx, _
                                                     ApproxSweepForInterval(Idx), ApproxRadiusForInterval(Idx), _
                                                     RollDir, REND_BEND_CURVE, STRMFG_DIRECTION
                        Next
                    End If
                    
                    If oNextCurve Is Nothing Then
                        MarkOtherEndOfCurve = True
                    ElseIf Not oNextArc Is Nothing Then
                        If SKIP_ROLL_REGION(oNextArc.Radius, oNextArc.SweepAngle) Then
                            ' Next curve is an arc that will be skipped so mark this region's roll boundary
                            MarkOtherEndOfCurve = True
                        Else
                            MarkOtherEndOfCurve = False
                        End If
                    Else
                        MarkOtherEndOfCurve = False
                    End If
                    
                    If MarkOtherEndOfCurve Then
                        If Multiplier < 0 Then
                            InflectStr = REND_BEND_RIGHT
                        Else
                            InflectStr = REND_BEND_LEFT
                        End If
                        
                        CreateRollBoundaryGeom2d oResourceManager, CurAccumLen, MarkBottomHeight, MarkTopHeight, _
                                                 oGeomCol2dOut, UpSide, "ROLLBOUNDARY-" & CStr(2 * i), _
                                                 ApproxSweepForInterval(UBound(ApproxSweepForInterval)), _
                                                 ApproxRadiusForInterval(UBound(ApproxRadiusForInterval)), _
                                                 RollDir, InflectStr, STRMFG_UNDEFINED_GEOMTYPE
                    End If
                    
                    ReDim BowStringDepths(0) As Double
                    ReDim ChordLengths(0) As Double
                    ReDim MidIntvlDists(0) As Double
                    ReDim FullIntvlDists(0) As Double
                    ReDim ApproxRadiusForInterval(0) As Double
                    ReDim ApproxSweepForInterval(0) As Double
                    
                End If
                
NextCurveInCS:
                AccumLen = CurAccumLen
                Set oCurve = Nothing
            Next ' Loop through curves in complex string
        End If ' Check if roll or bend info
    Next ' Loop through geom collection
    
    Set CreateProfileRollBoundaryMarks = oGeomCol2dOut

CleanUp:
    Set oMfgGeomCol2d = Nothing
    Set oProfileWrapper = Nothing
    Set oGeomCol2dOut = Nothing
    Set IntervalPoints = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function
