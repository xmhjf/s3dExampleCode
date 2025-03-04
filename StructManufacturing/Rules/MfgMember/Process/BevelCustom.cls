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
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: BevelCustom
'
' Description:  Provides the bevel deviation angle for the mfg profile
'
' Author:
'
' Comments:
' 02/26/2008    Teja    Created
'*******************************************************************************
Option Explicit

Private Const MODULE = "SKDY_MfgPrProcess.BevelCustom"

Implements IJDMfgProfileBevelRule
Implements IJDMfgProfileVaryingBevelRule

Private Function IJDMfgProfileBevelRule_GetBevelDeviationAngle(ByVal ProfilePart As Object) As Double
    Const METHOD = "BevelFixed: IJDMfgProfileBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler
        
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

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Sub

Private Sub IJDMfgProfileVaryingBevelRule_ManageBevelGaps(ByVal ProfilePart As GSCADMfgRulesDefinitions.IJProfilePart, ByVal GeomCollection As IMSCoreCollections.IJElements)
    Const METHOD = "IJDMfgProfileVaryingBevelRule_ManageBevelGaps"
    On Error GoTo ErrorHandler

        
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Sub

