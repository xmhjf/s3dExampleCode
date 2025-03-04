VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: BevelCustom
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
'
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgPlateProcess.BevelCustom"

Private Const CONNECTION_LENGTH = 0.003
Private Const CONNECTION_MIN_DIFF = 0.00005
Private Const CONNECTION_MAX_DIFF = 0.003
Private Const CONNECTION_ANGLE = 189.5

Private Const CONNECTION_VERTICAL_LINE = 1
Private Const CONNECTION_SLOPED_LINE = 2

Private Const UNDEFINED_BEVEL = 0
Private Const SQUARE_CUT = 1
Private Const Y_CUT = 2
Private Const X_CUT = 3
Private Const V_CUT = 4

Private Const AsRadians As Double = 1.74532925199433E-02
Private Const AsDegrees As Double = 57.2957795130823

Private Const M_PI As Double = 3.14159265358979

Implements IJDMfgPlateBevelRule
Implements IJDMfgPlateVaryingBevelRule

Private Function IJDMfgPlateBevelRule_GetBevelDeviationAngle(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "BevelFixed: IJDMfgPlateBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler

    ' Set the deviation angle when the physical connection should be splitted

    ' 3 degrees: GNEST hard-coded standard max. deviation-angle
    ' (angles are measured in radians in GSCAD)
    IJDMfgPlateBevelRule_GetBevelDeviationAngle = (3# * M_PI) / 180#

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Function

Private Sub IJDMfgPlateVaryingBevelRule_SplitMethodology(ByVal PlatePart As IJPlatePart, _
                                                         ByVal PortOfPlate As IJStructPort, _
                                                         ByVal Connection As IJAppConnection, _
                                                         HowToSplit As VaryingBevelSplitMethod, _
                                                         AngleValues() As Double, _
                                                         AllowableAbsoluteAngleError As Double, _
                                                         MinNumAnglesToSample As Long, _
                                                         SuggestSampleIntervalSpacing As Double, _
                                                         MinLengthOfSplitSegment As Double)

    Const METHOD = "IJDMfgPlateVaryingBevelRule_SplitMethodology"
    On Error GoTo ErrorHandler

    SuggestSampleIntervalSpacing = 0.5 ' Check attachment angles at every 0.5m of Phy Conn
    MinNumAnglesToSample = 8 ' Regardless of how many points are sampled per above suggestion,
                             ' ensure that at least "MinNumAnglesToSample" number of points
                             ' are evaluated for attachment angle.
    MinLengthOfSplitSegment = 0.0001 ' Ensure that points evaluated for attachment angles
                                     ' are at least "MinLengthOfSplitSegment" apart.
                                     
    ' METHODOLOGY 1:
    ' Split such that Angle Deviation between start and
    ' end of split region does not exceed specified value

'    HowToSplit = VaryingBevelSplitMethod.MaxAngleDeviation
'
'    ReDim AngleValues(1 To 1) As Double
'    AngleValues(1) = 3# ' In Degrees
'    AllowableAbsoluteAngleError = 0.01 ' In Degrees

    ' METHODOLOGY 2:
    ' Split PC whenever attachment angle is
    ' a multiple of specified offset from specified standard

    HowToSplit = VaryingBevelSplitMethod.OffsetFromStandard

    ReDim AngleValues(1 To 2) As Double
    AngleValues(1) = 90# ' In Degrees
    AngleValues(2) = 5#  ' In Degrees
    AllowableAbsoluteAngleError = 0.01 ' In Degrees


    ' METHODOLOGY 3:
    ' Split PC whenever attachment angle is one of specified values

'    HowToSplit = VaryingBevelSplitMethod.PresetBevelValues
'
'    ReDim AngleValues(1 To 33) As Double
'    ' WARNING: Preset values should cover entire range of possible attachment angles!
'    AllowableAbsoluteAngleError = 0.01 ' In Degrees
'    AngleValues(1) = 30
'    AngleValues(2) = 32
'    AngleValues(3) = 35
'    AngleValues(4) = 38
'    AngleValues(5) = 40
'    AngleValues(6) = 42
'    AngleValues(7) = 45
'    AngleValues(8) = 46
'    AngleValues(9) = 50
'    AngleValues(10) = 55
'    AngleValues(11) = 60
'    AngleValues(12) = 65
'    AngleValues(13) = 70
'    AngleValues(14) = 75
'    AngleValues(15) = 80
'    AngleValues(16) = 85
'    AngleValues(17) = 90
'    AngleValues(18) = 95
'    AngleValues(19) = 100
'    AngleValues(20) = 105
'    AngleValues(21) = 110
'    AngleValues(22) = 115
'    AngleValues(23) = 120
'    AngleValues(24) = 125
'    AngleValues(25) = 130
'    AngleValues(26) = 134
'    AngleValues(27) = 135
'    AngleValues(28) = 138
'    AngleValues(29) = 140
'    AngleValues(30) = 142
'    AngleValues(31) = 145
'    AngleValues(32) = 148
'    AngleValues(33) = 150

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Sub

Private Sub InitializeBevelDepthsAndAngles(ByRef BevelProps As IJMfgBevelDetailProperties)
    Const METHOD = "InitializeBevelDepthsAndAngles"
    On Error GoTo ErrorHandler

    BevelProps.Angle1_M = 0#
    BevelProps.Depth1_M = 0#
    BevelProps.Method1_M = 0
    
    BevelProps.Angle1_UM = 0#
    BevelProps.Depth1_UM = 0#
    BevelProps.Method1_UM = 0
    
    BevelProps.Angle2_M = 0#
    BevelProps.Depth2_M = 0#
    BevelProps.Method2_M = 0
    
    BevelProps.Angle2_UM = 0#
    BevelProps.Depth2_UM = 0#
    BevelProps.Method2_UM = 0

    BevelProps.Nose = 0#
    BevelProps.NoseAngle = 0#
    BevelProps.NoseMethod = 0
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Sub

Private Sub IJDMfgPlateVaryingBevelRule_SpecifyBevelProperties(ByVal PlatePart As IJPlatePart, _
                                                               ByVal PortOfPlate As IJStructPort, _
                                                               ByVal Connection As IJAppConnection, _
                                                               ByVal MfgOutput As IJMfgGeom2d, _
                                                               BevelProps As IJMfgBevel, _
                                                               vbControlBevelGaps As Boolean)

    Const METHOD = "IJDMfgPlateVaryingBevelRule_SpecifyBevelProperties"
    On Error GoTo ErrorHandler
    
    Dim bVaryingBevel As Boolean
    Dim AccessBevelParams As IJMfgBevelDetailProperties
    Set AccessBevelParams = BevelProps

    bVaryingBevel = False
    If (((AccessBevelParams.Depth1_M > 0.00001) And (AccessBevelParams.Method1_M > 0)) Or _
        ((AccessBevelParams.Depth1_UM > 0.00001) And (AccessBevelParams.Method1_UM > 0)) Or _
        ((AccessBevelParams.Depth2_M > 0.00001) And (AccessBevelParams.Method2_M > 0)) Or _
        ((AccessBevelParams.Depth2_UM > 0.00001) And (AccessBevelParams.Method2_UM > 0)) Or _
        ((AccessBevelParams.Nose > 0.00001) And (AccessBevelParams.NoseMethod > 0))) Then
        bVaryingBevel = True
        AccessBevelParams.VaryingBevel = True
    End If
    
    If bVaryingBevel = False Then
        vbControlBevelGaps = True
        Set AccessBevelParams = Nothing
        GoTo CleanUp
    End If

    Dim oPlate As IJPlate
    Set oPlate = PlatePart
    
    Dim StartAngle As Double
    StartAngle = BevelProps.RegionAttachmentAngle(StartOfRegionInPCVB)
    
    Dim EndAngle As Double
    EndAngle = BevelProps.RegionAttachmentAngle(EndOfRegionInPCVB)
    
    If Abs(StartAngle - 90 * AsRadians) > Abs(EndAngle - 90 * AsRadians) Then
        BevelProps.AttachmentAngle = StartAngle
    Else
        BevelProps.AttachmentAngle = EndAngle
    End If

    Dim RoundedAngle As Double
    
    If BevelProps.AttachmentAngle > 85 * AsRadians And _
       BevelProps.AttachmentAngle <= 90 * AsRadians Then

        InitializeBevelDepthsAndAngles AccessBevelParams

        AccessBevelParams.Depth1_M = 2# * oPlate.Thickness / 3#
        AccessBevelParams.Angle1_M = 5 * AsRadians

        AccessBevelParams.Depth1_UM = oPlate.Thickness / 3#
        AccessBevelParams.Angle1_UM = 40 * AsRadians

    ElseIf BevelProps.AttachmentAngle > 90 * AsRadians And _
           BevelProps.AttachmentAngle <= 95 * AsRadians Then

        InitializeBevelDepthsAndAngles AccessBevelParams

        AccessBevelParams.Depth1_UM = 2# * oPlate.Thickness / 3#
        AccessBevelParams.Angle1_UM = 5 * AsRadians

        AccessBevelParams.Depth1_M = oPlate.Thickness / 3#
        AccessBevelParams.Angle1_M = 40 * AsRadians

    ElseIf BevelProps.AttachmentAngle > 80 * AsRadians And _
           BevelProps.AttachmentAngle <= 85 * AsRadians Then

        InitializeBevelDepthsAndAngles AccessBevelParams

        AccessBevelParams.Depth1_M = 2# * oPlate.Thickness / 3#
        AccessBevelParams.Angle1_M = 10 * AsRadians

        AccessBevelParams.Depth1_UM = oPlate.Thickness / 3#
        AccessBevelParams.Angle1_UM = 35 * AsRadians

    ElseIf BevelProps.AttachmentAngle > 95 * AsRadians And _
           BevelProps.AttachmentAngle <= 100 * AsRadians Then

        InitializeBevelDepthsAndAngles AccessBevelParams

        AccessBevelParams.Depth1_UM = 2# * oPlate.Thickness / 3#
        AccessBevelParams.Angle1_UM = 10 * AsRadians

        AccessBevelParams.Depth1_M = oPlate.Thickness / 3#
        AccessBevelParams.Angle1_M = 35 * AsRadians

    ElseIf BevelProps.AttachmentAngle > 100 * AsRadians Then

        InitializeBevelDepthsAndAngles AccessBevelParams
        
        RoundedAngle = Round((BevelProps.AttachmentAngle - 100 * AsRadians) / (5 * AsRadians))
        RoundedAngle = RoundedAngle * 5 * AsRadians + 10 * AsRadians

        AccessBevelParams.Depth1_UM = oPlate.Thickness
        AccessBevelParams.Angle1_UM = RoundedAngle

    ElseIf BevelProps.AttachmentAngle < 80 * AsRadians Then

        InitializeBevelDepthsAndAngles AccessBevelParams
        
        RoundedAngle = Round((80 * AsRadians - BevelProps.AttachmentAngle) / (5 * AsRadians))
        RoundedAngle = 10 * AsRadians + RoundedAngle * 5 * AsRadians

        AccessBevelParams.Depth1_M = oPlate.Thickness
        AccessBevelParams.Angle1_M = RoundedAngle

    End If

    vbControlBevelGaps = True
    
CleanUp:
    Set oPlate = Nothing
    Set AccessBevelParams = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJDMfgPlateVaryingBevelRule_ManageBevelGaps(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart, ByVal GeomCollection As IMSCoreCollections.IJElements)
    Const METHOD = "IJDMfgPlateVaryingBevelRule_ManageBevelGaps"
    On Error GoTo ErrorHandler
    Dim ind As Long
    Dim lGeomCount As Long
    Dim dOffsetDiff As Double
    Dim oThisGeom2d As IJMfgGeom2d
    Dim oPrevGeom2d As IJMfgGeom2d
    Dim oThisBevel As IJMfgBevelDetailProperties
    Dim oPrevBevel As IJMfgBevelDetailProperties
    Dim oMfgMatHGeom As New MfgMathGeom
    
    lGeomCount = GeomCollection.Count
    
    For ind = 1 To lGeomCount
        ' Get this geom2d and previous geom2d
        Set oThisBevel = Nothing
        Set oThisGeom2d = GeomCollection.Item(ind)
        Set oThisBevel = oThisGeom2d.GetBevel
        
        ' Get the bevels also
        If ind = 1 Then
            Set oPrevGeom2d = GeomCollection.Item(lGeomCount)
            Set oPrevBevel = oPrevGeom2d.GetBevel
        End If
        
        If ((Not oThisBevel Is Nothing) And (Not oPrevBevel Is Nothing)) Then
            ' Get the difference of bevel corrections
            dOffsetDiff = Abs(oThisBevel.CalculatedBevelOffset - oPrevBevel.CalculatedBevelOffset)
            ' Process the code only if there is a considerable difference in the bevel correction values
            If dOffsetDiff > 0.000001 Then
                Dim dStartParam As Double, dEndParam As Double, dDummy As Double
                Dim dX As Double, dY As Double, dZ As Double, dAngle As Double
                Dim oPrevCSCurve As IJCurve, oThisCSCurve As IJCurve
                Dim oPrevTangent As New DVector
                Dim oThisTangent As New DVector
                Dim oMaxOffsetGeom2d As IJMfgGeom2d
                Dim oMinOffsetGeom2d As IJMfgGeom2d
                Dim oZAxis As New DVector
                Dim VToVChange As Boolean
                Dim bVaryingBevel As Boolean
                Dim bBevelTypeChange As Boolean
                Dim dVAngle1 As Double, dVAngle2 As Double
                
                VToVChange = False
                bVaryingBevel = False
                bBevelTypeChange = True
                
                ' Evaluate the tangent vector at the end point of previous curve
                Set oPrevCSCurve = oPrevGeom2d.GetGeometry
                oPrevCSCurve.ParamRange dStartParam, dEndParam
                oPrevCSCurve.Evaluate dEndParam, dDummy, dDummy, dDummy, dX, dY, dZ, dDummy, dDummy, dDummy
                oPrevTangent.Set -1# * dX, -1# * dY, -1# * dZ
                
                ' Evaluate the tangent vector at the start point of this curve
                Set oThisCSCurve = oThisGeom2d.GetGeometry
                oThisCSCurve.ParamRange dStartParam, dEndParam
                oThisCSCurve.Evaluate dEndParam, dDummy, dDummy, dDummy, dX, dY, dZ, dDummy, dDummy, dDummy
                oThisTangent.Set dX, dY, dZ
                
                oZAxis.Set 0#, 0#, 1#
                
                ' Evaluate the angle between the previous tangent and this tangent
                dAngle = AsDegrees * oPrevTangent.Angle(oThisTangent, oZAxis)
                
                ' Check if this is a V-Cut to V-Cut change
                If (GetBevelType(oThisBevel, dVAngle1) = V_CUT And GetBevelType(oPrevBevel, dVAngle2) = V_CUT) Then
                    VToVChange = True
                End If
                
                If (GetBevelType(oThisBevel, dVAngle1) = GetBevelType(oPrevBevel, dVAngle2)) Then
                    bBevelTypeChange = False
                End If
                
                ' Check if this is a Varying Bevel case or Constant bevel case
                If (oThisBevel.VaryingBevel And oPrevBevel.VaryingBevel) Then
                    bVaryingBevel = True
                End If
                
                ' Check which has Geom2d has more offset
                If oThisBevel.CalculatedBevelOffset > oPrevBevel.CalculatedBevelOffset Then
                    Set oMaxOffsetGeom2d = oThisGeom2d
                    Set oMinOffsetGeom2d = oPrevGeom2d
                Else
                    Set oMaxOffsetGeom2d = oPrevGeom2d
                    Set oMinOffsetGeom2d = oThisGeom2d
                End If
                
                ' If it is a varying bevel, we should always close with 3mm sloped segment
                If (bVaryingBevel) Then
                    If VToVChange Then
                        If (((dVAngle1 > 0.0001) And (dVAngle2 < 0.0001)) Or _
                            ((dVAngle1 < 0.0001) And (dVAngle2 > 0.0001))) Then
                            oMfgMatHGeom.CloseGapBetweenGeometries oMaxOffsetGeom2d, oMinOffsetGeom2d, False, False, CONNECTION_SLOPED_LINE, CONNECTION_LENGTH
                        End If
                    ElseIf bBevelTypeChange Then
                        oMfgMatHGeom.CloseGapBetweenGeometries oMaxOffsetGeom2d, oMinOffsetGeom2d, False, False, CONNECTION_SLOPED_LINE, CONNECTION_LENGTH
                    End If
                Else
                    ' If it is a Constant bevel and if the bevel difference is between 3mm and 0.05 mm,
                    ' we should  close with 3mm sloped segment
                    If ((dOffsetDiff > CONNECTION_MIN_DIFF) And (dOffsetDiff < CONNECTION_MAX_DIFF)) Then
                        oMfgMatHGeom.CloseGapBetweenGeometries oMaxOffsetGeom2d, oMinOffsetGeom2d, False, False, CONNECTION_SLOPED_LINE, CONNECTION_LENGTH
                    ElseIf dOffsetDiff > CONNECTION_MAX_DIFF Then
                    ' If the bevel difference is more than 3mm, connect with a 3mm linear segment perpendicular to the maximum edge
                        oMfgMatHGeom.CloseGapBetweenGeometries oMaxOffsetGeom2d, oMinOffsetGeom2d, True, True, CONNECTION_VERTICAL_LINE
                    ElseIf ((VToVChange = True) And (dAngle > CONNECTION_ANGLE)) Then
                    ' If it is a V-Cut to V-Cut change and if the angle between adajacent curve is more than 189.5 degrees, connect with a sloped segment of length 3mm
                        oMfgMatHGeom.CloseGapBetweenGeometries oMaxOffsetGeom2d, oMinOffsetGeom2d, False, False, CONNECTION_SLOPED_LINE, CONNECTION_LENGTH
                    End If
                End If
            End If
        End If
        
        Set oPrevGeom2d = oThisGeom2d
        Set oPrevBevel = oThisBevel
    Next ind
        
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
  
End Sub

Private Function GetBevelType(ByVal oMfgBevel As IJMfgBevelDetailProperties, ByRef dVAngle As Double) As Long
    Const METHOD = "GetBevelType"
    On Error GoTo ErrorHandler

    Dim bChamferDepth_M As Boolean
    Dim bChamferDepth_UM As Boolean
    Dim bDepth1_M As Boolean
    Dim bDepth1_UM As Boolean
    Dim bDepth2_M As Boolean
    Dim bDepth2_UM As Boolean
    Dim bNoseDepth As Boolean
    Dim bSqaureCut As Boolean
    Dim bVCut1 As Boolean
    Dim bVCut2 As Boolean
    Dim dAngle As Double
    
    bVCut1 = False
    bVCut2 = False
    bChamferDepth_M = (Abs(oMfgBevel.ChamferDepth_M) > 0.0001)
    bChamferDepth_UM = (Abs(oMfgBevel.ChamferDepth_UM) > 0.0001)
    bDepth1_M = (Abs(oMfgBevel.Depth1_M) > 0.0001)
    bDepth1_UM = (Abs(oMfgBevel.Depth1_UM) > 0.0001)
    bDepth2_M = (Abs(oMfgBevel.Depth2_M) > 0.0001)
    bDepth2_UM = (Abs(oMfgBevel.Depth2_UM) > 0.0001)
    bNoseDepth = (Abs(oMfgBevel.Nose) > 0.0001)
    
    dVAngle = 0#
    
    If bChamferDepth_M Then
        If (Abs(oMfgBevel.ChamferAngle_M) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            bVCut1 = True
        End If
    End If
    
    If bChamferDepth_UM Then
        If (Abs(oMfgBevel.ChamferAngle_UM) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            bVCut1 = True
        End If
    End If
    
    If bDepth1_M Then
        dAngle = oMfgBevel.Angle1_M
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bDepth1_UM Then
        dAngle = oMfgBevel.Angle1_UM
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bDepth2_M Then
        dAngle = oMfgBevel.Angle2_M
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bDepth2_UM Then
        dAngle = oMfgBevel.Angle2_UM
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bNoseDepth Then
        dAngle = oMfgBevel.NoseAngle
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bSqaureCut Then
        If bVCut1 Or bVCut2 Then
            GetBevelType = Y_CUT
        Else
            GetBevelType = SQUARE_CUT
        End If
    ElseIf bVCut1 And bVCut2 Then
        GetBevelType = X_CUT
    ElseIf bVCut1 Or bVCut2 Then
        GetBevelType = V_CUT
    Else
        GetBevelType = UNDEFINED_BEVEL
    End If
        
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Function

