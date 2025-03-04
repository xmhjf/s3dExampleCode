VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NACalcStandard"
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

    Dim bstrSectionType As String
    bstrSectionType = oProfileClass.SectionType
    
    oProfileClass.Get_Centroid pXValue, pYValue
                    
    Dim bstrSectionName As String
    bstrSectionName = oProfileClass.SectionName
    
    Dim oQueryOutputValues() As Variant
    Dim strQuery As String
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oProfilePartHlpr As MfgRuleHelpers.ProfilePartHlpr
    
    Set oProfilePartHlpr = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHlpr.object = oProfilePart
                        
    If oProfilePartHlpr.ProfileHasMfgPart(oMfgProfilePart) Then
        strQuery = "SELECT ConvexCentroidX, ConvexCentroidY, ConcaveCentroidX, ConcaveCentroidY From JUAMfgProfileNAxisAB WHERE (SectionName = '" + bstrSectionName + "') AND (Machine = 'ABB')"
        
        On Error Resume Next
        oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
        If (UBound(oQueryOutputValues) > 0) Then
            oMfgProfilePart.SetNeutralAxisValues oQueryOutputValues(0), oQueryOutputValues(1), oQueryOutputValues(2), oQueryOutputValues(3)
        End If
    End If
    Set oMfgProfilePart = Nothing
    Set oMfgCatalogQueryHelper = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2030, , "RULES")
End Sub

