VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavalArchLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: NavalArchLines
'
' Description:  Example to show the naval architecture lines on the pinjig
'
' Author: Anand Hariharan
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.NavalArchLines"

Implements IJDMfgPinJigRulePerRemarkingType
Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IMSCoreCollections.IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim RefCurveDataColl As Collection
    Set RefCurveDataColl = GetRefCurveGeomDataFromPinJigSupportedPlates(PartialPinJig)
    
    Dim ReturnColl As IJElements
    Set ReturnColl = New JObjectCollection
    
    Dim oRefCurveData As IJRefCurveData
    For Each oRefCurveData In RefCurveDataColl
        ReturnColl.Add oRefCurveData.ParentReferenceCurve
    Next
    
    GetAPSMarkingLines PartialPinJig, STRMFG_NAVALARCHLINE, ReturnColl
    
    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = ReturnColl
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5016, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IMSCoreCollections.IJElements) As String
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsNavalArchLine]"
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IMSCoreCollections.IJElements) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler
    
    Dim ReturnColl As IJMfgGeomCol3d
    
    CreateNavalArchMarks PartialPinJig, ElemsToRemark, ReturnColl
    
    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl

    Set ReturnColl = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5017, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
