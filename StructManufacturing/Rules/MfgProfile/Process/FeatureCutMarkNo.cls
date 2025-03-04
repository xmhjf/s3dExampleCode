VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeatureCutMarkNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: FeatureCutMarkNo
'
' Description:  To customize the cut/mark/ignore  options for feature cuts.
'  History:
'       StructMfg         September 15th, 2010   created
'
'*******************************************************************************
Option Explicit
Const MODULE = "MfgProfileProcess.FeatureCutMarkNo"
Implements IJMfgFeatureCutMark

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
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2057, , "RULES")
    GoTo CleanUp
End Function

