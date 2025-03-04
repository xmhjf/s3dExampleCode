VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NACalcBuiltUp"
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
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileNeutralAxisRule

Const MODULE = "MfgProfileProcess"

Private Sub IJDMfgProfileNeutralAxisRule_GetNeutralAxis(ByVal oProfilePart As Object, pXValue As Double, pYValue As Double)
    Const METHOD = "NeutralAxisCalculated: IJDMfgProfileNeutralAxisRule_GetNeutralAxis"
    On Error GoTo ErrorHandler
    
    ' Use the profile part as input
    Dim oProfileClass As StructDetailObjects.ProfilePart
    
    Set oProfileClass = New StructDetailObjects.ProfilePart
    
    Set oProfileClass.object = oProfilePart

    If oProfileClass.IsCrossSectionABuiltUp = False Then
        oProfileClass.Get_Centroid pXValue, pYValue
        Exit Sub
    End If
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oProfilePartHlpr As MfgRuleHelpers.ProfilePartHlpr
    
    Set oProfilePartHlpr = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHlpr.object = oProfilePart
                        
    If Not (oProfilePartHlpr.ProfileHasMfgPart(oMfgProfilePart)) Then
        Exit Sub
    End If
    
    Set oProfilePartHlpr = Nothing
    
    Dim A_WL As Double, t1_WT As Double
    Dim B_FL As Double, t2_FT As Double, dComputedNAValue As Double
    
    A_WL = oProfileClass.WebLength
    t1_WT = oProfileClass.WebThickness
    
    B_FL = oProfileClass.FlangeLength
    t2_FT = oProfileClass.FlangeThickness
    
    Dim oQueryOutputValues() As Variant
    Dim strQuery As String
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
    
    strQuery = "SELECT CentroidX, CentroidY From JUAMfgProfileNAxisBU WHERE (WebDepth = " + CStr(A_WL) + ") AND (WebThickness = " + CStr(t1_WT) + ") AND (FlangeWidth = " + CStr(B_FL) + ") AND (FlangeThickness = " + CStr(t2_FT) + ") AND (Machine = 'TABR')"
    
    On Error Resume Next
    oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
        
    If (UBound(oQueryOutputValues) > 0) Then
        oMfgProfilePart.SetNeutralAxisValues oQueryOutputValues(0), oQueryOutputValues(1), oQueryOutputValues(0), oQueryOutputValues(1)
    Else
        Dim dNumerator As Double, dDenominator As Double
       
        dNumerator = ((B_FL) * ((A_WL + t2_FT) ^ 2)) - ((A_WL ^ 2) * (B_FL - t1_WT))
        dDenominator = 2# * ((B_FL * (A_WL + t2_FT)) - A_WL * (B_FL - t1_WT))
        dComputedNAValue = dNumerator / dDenominator
        
        oMfgProfilePart.SetNeutralAxisValues 0#, dComputedNAValue, 0#, dComputedNAValue
    End If
    
    Set oMfgProfilePart = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2030, , "RULES")
End Sub
