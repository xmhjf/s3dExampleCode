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
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: BevelLookup
'
' Description:  Provides the bevel deviation angle for the mfg profile
'
' Author:
'
' Comments:
' 02/26/2008    Teja    Created
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgProfileProcess.BevelLookup"

Implements IJDMfgProfileBevelRule
Implements IJDMfgProfileVaryingBevelRule

Private Function IJDMfgProfileBevelRule_GetBevelDeviationAngle(ByVal ProfilePart As Object) As Double
    Const METHOD = "BevelFixed: IJDMfgProfileBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler

    ' Set the deviation angle when the physical connection should be splitted
    
    ' 3 degrees: GNEST hard-coded standard max. deviation-angle
    ' (angles are measured in radians in GSCAD)
    IJDMfgProfileBevelRule_GetBevelDeviationAngle = (3# * 3.141529) / 180#
        
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2041, , "RULES")
End Function

Private Sub IJDMfgProfileVaryingBevelRule_SplitMethodology(ByVal ProfilePart As IJProfilePart, _
                                                           ByVal PortOfProfile As IJStructPort, _
                                                           ByVal Connection As IJAppConnection, _
                                                           HowToSplit As VaryingBevelSplitMethod, _
                                                           AngleValues() As Double, _
                                                           AllowableAbsoluteAngleError As Double, _
                                                           MinNumAnglesToSample As Long, _
                                                           SuggestSampleIntervalSpacing As Double, _
                                                           MinLengthOfSplitSegment As Double)

    Const METHOD = "IJDMfgProfileVaryingBevelRule_SplitMethodology"
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


Private Sub IJDMfgProfileVaryingBevelRule_SpecifyBevelProperties(ByVal ProfilePart As IJProfilePart, _
                                                                 ByVal PortOfProfile As IJStructPort, _
                                                                 ByVal Connection As IJAppConnection, _
                                                                 ByVal MfgOutput As IJMfgGeom2d, _
                                                                 BevelProps As IJMfgBevel, _
                                                                 vbControlBevelGaps As Boolean)

    Const METHOD = "IJDMfgPlateVaryingBevelRule_SpecifyBevelProperties"
    On Error GoTo ErrorHandler

    Const AsRadians As Double = 1.74532925199433E-02
    
    Dim oProfileSection As IJDProfileSection
    Set oProfileSection = ProfilePart
    
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = oProfileSection.CrossSection
    
    Dim oCSPartClass As IJDCrossSectionPartClass
    Set oCSPartClass = oCrossSection.GetPartClass
    
    Dim SectionName As String
    SectionName = oCSPartClass.CrossSectionTypeName
    
    If (UCase(SectionName) = "EQUALANGLE" Or _
        UCase(SectionName) = "UNEQUALANGLE") And _
       (PortOfProfile.ContextID = CTX_BASE Or _
        PortOfProfile.ContextID = CTX_OFFSET) _
    Then
        BevelProps.Purpose = 1 ' Show bevel but not apply offset
    End If
        
CleanUp:
    Set oProfileSection = Nothing
    Set oCrossSection = Nothing
    Set oCSPartClass = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJDMfgProfileVaryingBevelRule_ManageBevelGaps(ByVal ProfilePart As GSCADMfgRulesDefinitions.IJProfilePart, ByVal GeomCollection As IMSCoreCollections.IJElements)

End Sub



