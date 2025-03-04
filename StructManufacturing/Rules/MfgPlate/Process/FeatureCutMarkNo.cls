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
' Project: MfgPlateProcess
' Module: FeatureCutMarkNo
'
' Description:  To customize the cut/mark/ignore  options for feature cuts.
'  History:
'       StructMfg         September 15th, 2010   created
'
'*******************************************************************************
Option Explicit
Implements IJMfgFeatureCutMark
Implements IJMfgFeatureBridge
Const MODULE = "MfgPlateProcess.FeatureCutMarkNo"

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
        If Not oStructCutout Is Nothing Then
            If oStructCutout.CutoutProductionMethod = cpmMark Then
                IJMfgFeatureCutMark_GetFeatureCutMarkOption = MarkFeature
            ElseIf oStructCutout.CutoutProductionMethod = cpmBridge Then
                IJMfgFeatureCutMark_GetFeatureCutMarkOption = BridgeFeature
            End If
        End If
        Set oStructCutout = Nothing
        Set oHoleFeature = Nothing
     ElseIf TypeOf oFeatureObj Is IJStructCutout Then
        Set oStructCutout = oFeatureObj
        If oStructCutout.CutoutProductionMethod = cpmMark Then
            IJMfgFeatureCutMark_GetFeatureCutMarkOption = MarkFeature
        ElseIf oStructCutout.CutoutProductionMethod = cpmBridge Then
            IJMfgFeatureCutMark_GetFeatureCutMarkOption = BridgeFeature
        End If
        Set oStructCutout = Nothing
    End If
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1075, , "RULES")
    GoTo CleanUp
End Function

Private Function IJMfgFeatureBridge_GetBridgeNamesList(ByVal oMfgDispObj As Object, ByVal oFeatureObj As Object) As String()
    Dim BridgeNamesList() As String
    Dim i As Integer, lBridgeIdx
    Dim strQuery As String
    Dim strTypeQuery As String
    Dim strBridgeDefaultOption As String

    Dim bIsPartOfConour As Boolean
    Dim dOpeningDistance As Double
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper

    On Error Resume Next
    oMfgRuleHelper.GetOpeningDistanceFromContour oMfgDispObj, oFeatureObj, bIsPartOfConour, dOpeningDistance
    If bIsPartOfConour Then
        strBridgeDefaultOption = "Bridge_o0_w150"
    ElseIf dOpeningDistance < 0.15 Then
        strBridgeDefaultOption = "Bridge_o150_w150"
    ElseIf dOpeningDistance < 0.3 Then
        strBridgeDefaultOption = "Bridge_o0_w100"
    Else
        strBridgeDefaultOption = "Bridge_o-1_w150"
    End If
    
    ' BP.Type 1 for Holes, 2 for Features and 3 for all
    If TypeOf oFeatureObj Is IJOpeningEntity Then
        strTypeQuery = "where BP.Type <> 2"
    Else
        strTypeQuery = "where BP.Type <> 1"
    End If

    strQuery = "Select CI.Name from JUASMPlateBridgeParams  BP join REFDATClassItem CI ON BP.oid=CI.OID " & strTypeQuery

    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper

    Dim oQueryOutputValues() As Variant
    oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)

    If (UBound(oQueryOutputValues) > LBound(oQueryOutputValues)) Then
        ReDim BridgeNamesList(UBound(oQueryOutputValues)) As String
        lBridgeIdx = 0
        For i = LBound(oQueryOutputValues) To UBound(oQueryOutputValues)
            BridgeNamesList(lBridgeIdx) = CStr(oQueryOutputValues(i))
            If (BridgeNamesList(lBridgeIdx) = strBridgeDefaultOption) Then
                Dim strTemp As String
                strTemp = BridgeNamesList(0)
                BridgeNamesList(0) = strBridgeDefaultOption
                BridgeNamesList(lBridgeIdx) = strTemp
            End If
            lBridgeIdx = lBridgeIdx + 1
        Next
    End If

    IJMfgFeatureBridge_GetBridgeNamesList = BridgeNamesList

    Set oMfgCatalogQueryHelper = Nothing
    Set oMfgRuleHelper = Nothing
End Function
