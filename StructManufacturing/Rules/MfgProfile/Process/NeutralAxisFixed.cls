VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NeutralAxisFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module:
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileNeutralAxisRule

Const MODULE = "MfgProfileProcess"

Private Sub IJDMfgProfileNeutralAxisRule_GetNeutralAxis(ByVal oProfilePart As Object, pXValue As Double, pYValue As Double)
    Const METHOD = "NeutralAxisFixed: IJDMfgProfileNeutralAxisRule_GetNeutralAxis"
    On Error GoTo ErrorHandler
    
    ' Use the profile part as input
    Dim oProfileClass As Object
    If TypeOf oProfilePart Is IJStiffenerPart Then
        Set oProfileClass = New StructDetailObjects.ProfilePart
    ElseIf TypeOf oProfilePart Is IJBeamPart Then
        Set oProfileClass = New StructDetailObjects.BeamPart
    End If
    Set oProfileClass.object = oProfilePart

   'Initialize the values by checking the values from reference data.
    oProfileClass.Get_Centroid pXValue, pYValue

    Dim oMFGRuleHelper As MfgRuleHelpers.Helper
    Set oMFGRuleHelper = New MfgRuleHelpers.Helper

    Dim oMfgRuleProfilePart As MfgRuleHelpers.ProfilePartHlpr
    Set oMfgRuleProfilePart = New MfgRuleHelpers.ProfilePartHlpr
    Set oMfgRuleProfilePart.object = oProfilePart
    Dim oMfgProfilePart As IJMfgProfilePart
    
    If oMfgRuleProfilePart.ProfileHasMfgPart(oMfgProfilePart) = True Then
        Dim rulevalue As String

        rulevalue = oMFGRuleHelper.GetRuleSettings(oMfgProfilePart, PROCESS_SETTINGS, "ProfileUnfold")

        If rulevalue = "Ignore" Then        ' 2149 = "Ignore"
        
        ' check if the cross-section is a built-up or a flatbar
            Dim bIsBuiltUp As Boolean
            bIsBuiltUp = oProfileClass.IsCrossSectionABuiltUp
            
            If bIsBuiltUp = True Then
                
                Dim Pr_Curvature As ProfileCurvature
                Pr_Curvature = oMfgRuleProfilePart.CurvatureType
                
                If Pr_Curvature <> PROFILE_CURVATURE_Straight Then
                     
                    If Pr_Curvature = PROFILE_CURVATURE_CurvedAlongFlange Or _
                       Pr_Curvature = (PROFILE_CURVATURE_CurvedAlongFlange Or PROFILE_CURVATURE_KnuckledAlongFlange) Or _
                       Pr_Curvature = (PROFILE_CURVATURE_CurvedAlongFlange Or PROFILE_CURVATURE_CurvedAlongWeb) Then
                    
                        pYValue = oProfileClass.FlangeThickness / 2
                    Else
                        If Pr_Curvature = PROFILE_CURVATURE_CurvedAlongWeb Or _
                           Pr_Curvature = (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_KnuckledAlongWeb) Then
                            pYValue = oProfileClass.FlangeThickness + oProfileClass.WebLength / 2
                        End If
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2031, , "RULES")
End Sub
 
