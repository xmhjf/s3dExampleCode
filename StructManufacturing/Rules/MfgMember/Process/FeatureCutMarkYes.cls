VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeatureCutMarkYes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: FeatureCutMarkYes
'
' Description:  To customize the cut/mark/ignore  options for feature cuts.
'  History:
'       StructMfg         September 15th, 2010   created
'
'*******************************************************************************
Implements IJMfgFeatureCutMark
Const MODULE = "MfgMemberProcess.FeatureCutMarkYes"
Private Const MAX_FEATURE_SIZE_FOR_MARK As Double = 0.035
Option Explicit

Private Function IJMfgFeatureCutMark_GetFeatureCutMarkOption( _
        ByVal oMfgDispObj As Object, _
        ByVal oFeatureObj As Object) As GSCADMfgRulesDefinitions.EnumStrMfgFeatureCutMark
    ' oMfgDispObj - Input - Manufacturing Plate Object
    ' oFeatureObj - Input - Feature object on the detailed plate part
    
    Const METHOD = "FeatureCutMark: IJMfgFeatureCutMark_GetFeatureCutMarkOption"
    On Error GoTo ErrorHandler
    
    IJMfgFeatureCutMark_GetFeatureCutMarkOption = CutFeature
    
    If TypeOf oFeatureObj Is IJHoleFeature Then
        Dim oHoleFeature As IJHoleFeature
        Dim oStructCutout As IJStructCutout
        
        Set oHoleFeature = oFeatureObj
        Set oStructCutout = oHoleFeature.HoleTrace
        If oStructCutout.CutoutProductionMethod = cpmMark Then
            IJMfgFeatureCutMark_GetFeatureCutMarkOption = MarkFeature
        End If
        Set oStructCutout = Nothing
        Set oHoleFeature = Nothing
     ElseIf TypeOf oFeatureObj Is IJStructCutout Then
        Set oStructCutout = oFeatureObj
        If oStructCutout.CutoutProductionMethod = cpmMark Then
            IJMfgFeatureCutMark_GetFeatureCutMarkOption = MarkFeature
        End If
        Set oStructCutout = Nothing
    End If
    
    ' Check the Feature size and set the FeatureCutMarkOption accordingly
    If Not TypeOf oFeatureObj Is IJOpeningEntity Then
        If TypeOf oFeatureObj Is IJStructFeature Then
            Dim oStructFeature As IJStructFeature
            Set oStructFeature = oFeatureObj
            If ((oStructFeature.get_StructFeatureType = SF_WebCut) Or _
                (oStructFeature.get_StructFeatureType = SF_FlangeCut)) Then
                IJMfgFeatureCutMark_GetFeatureCutMarkOption = CutFeature
                Set oStructFeature = Nothing
                GoTo Cleanup
            End If
            Set oStructFeature = Nothing
        End If
        
        Dim bMarkFeature    As Boolean
        'bMarkFeature = IsFeatureRangeWithInMaxSize(oMfgDispObj, oFeatureObj, MAX_FEATURE_SIZE_FOR_MARK)
        
        If bMarkFeature = True Then
            IJMfgFeatureCutMark_GetFeatureCutMarkOption = MarkFeature
        End If
    End If
    
Cleanup:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2057, , "RULES")
    GoTo Cleanup
End Function

