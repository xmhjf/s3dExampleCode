VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: BevelLookup
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
'
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgPlateProcess.BevelLookup"

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

    vbControlBevelGaps = False
    
CleanUp:
    Set oPlate = Nothing
    Set AccessBevelParams = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJDMfgPlateVaryingBevelRule_ManageBevelGaps(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart, ByVal GeomCollection As IMSCoreCollections.IJElements)

End Sub
