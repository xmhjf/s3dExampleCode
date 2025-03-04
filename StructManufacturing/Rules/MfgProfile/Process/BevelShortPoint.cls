VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelShortPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: BevelShortPoint
'
' Description:  Accounts for bevel Butt-PCs at Profile ends to the short-point
'
' Author: Nautilus
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgProfileProcess.BevelShortPoint::"

Implements IJDMfgProfileBevelRule
Implements IJDMfgProfileVaryingBevelRule

Private Function IJDMfgProfileBevelRule_GetBevelDeviationAngle(ByVal ProfilePart As Object) As Double
    Const METHOD = "BevelFixed: IJDMfgProfileBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler

    ' Set the deviation angle when the physical connection should be splitted
    
    ' 3 degrees: GNEST hard-coded standard max. deviation-angle
    ' (angles are measured in radians in GSCAD)
    IJDMfgProfileBevelRule_GetBevelDeviationAngle = (3# * 3.14159265358979) / 180#
        
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

    ' METHODOLOGY:
    ' Split such that Angle Deviation between start and
    ' end of split region does not exceed specified value

    HowToSplit = VaryingBevelSplitMethod.MaxAngleDeviation

    ReDim AngleValues(1 To 1) As Double
    AngleValues(1) = 3# ' In Degrees
    AllowableAbsoluteAngleError = 0.01 ' In Degrees

    SuggestSampleIntervalSpacing = 0.5 ' Check attachment angles at every 0.5m of Phy Conn
    MinNumAnglesToSample = 8 ' Regardless of how many points are sampled per above suggestion,
                             ' ensure that at least "MinNumAnglesToSample" number of points
                             ' are evaluated for attachment angle.
    MinLengthOfSplitSegment = 0.0001 ' Ensure that points evaluated for attachment angles
                                     ' are at least "MinLengthOfSplitSegment" apart.

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

    Const METHOD = "IJDMfgProfileVaryingBevelRule_SpecifyBevelProperties"
    On Error GoTo ErrorHandler

    ' Check if the port is one of profile's ends (i.e., base or offset)
    Dim portContext As eUSER_CTX_FLAGS
    portContext = PortOfProfile.ContextID
    If portContext <> CTX_BASE And portContext <> CTX_OFFSET Then
        GoTo CleanUp
    End If
    
    ' Check whether the PC is of Butt Type
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = ProfilePart

    Dim bIsCrossOfTee As Boolean
    Dim eConnType As ContourConnectionType
    oSDPartSupport.GetConnectionTypeForContour Connection, _
                                               eConnType, _
                                               bIsCrossOfTee
                                               
    If eConnType <> PARTSUPPORT_CONNTYPE_BUTT Then
        GoTo CleanUp
    End If
        
    ' Set Bevel to Short Point
    Dim oBevelDetails As IJMfgBevelDetailProperties
    Set oBevelDetails = BevelProps
    
    ' New property made available by CR 316897
    oBevelDetails.MfgMaxShapePosition = MfgMaxShapePos_ShortPoint
    Set oBevelDetails = Nothing
        
CleanUp:
    Set oSDPartSupport = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJDMfgProfileVaryingBevelRule_ManageBevelGaps(ByVal ProfilePart As GSCADMfgRulesDefinitions.IJProfilePart, ByVal GeomCollection As IMSCoreCollections.IJElements)

End Sub



