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
' Project: MfgMemberProcess
' Module: FeatureCutMarkNo
'
' Description:  To customize the cut/mark/ignore  options for feature cuts.
'  History:
'       StructMfg         September 15th, 2010   created
'
'*******************************************************************************
Implements IJMfgFeatureCutMark
Const MODULE = "MfgMemberProcess.FeatureCutMarkNo"
Option Explicit
Private Function IJMfgFeatureCutMark_GetFeatureCutMarkOption( _
        ByVal oMfgDispObj As Object, _
        ByVal oFeatureObj As Object) As GSCADMfgRulesDefinitions.EnumStrMfgFeatureCutMark
    ' oMfgDispObj - Input - Manufacturing Plate Object
    ' oFeatureObj - Input - Feature object on the detailed plate part
    
    Const METHOD = "FeatureCutMark: IJMfgFeatureCutMark_GetFeatureCutMarkOption"
    On Error GoTo ErrorHandler
    
    IJMfgFeatureCutMark_GetFeatureCutMarkOption = CutFeature
    
    If TypeOf oFeatureObj Is IJStructCutout Then
        Dim oStructCutout As IJStructCutout
                Set oStructCutout = oFeatureObj
        If oStructCutout.CutoutProductionMethod = cpmMark Then
            IJMfgFeatureCutMark_GetFeatureCutMarkOption = MarkFeature
        End If
        Set oStructCutout = Nothing
    End If
    
Cleanup:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2057, , "RULES")
    GoTo Cleanup
End Function

