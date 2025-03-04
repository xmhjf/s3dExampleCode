VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module:
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2010.06.22    Siva    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const Module = "MfgPlateProcess"

Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "MarginFixed: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error Resume Next
    Dim oPortColl As IJElements
    Dim oFlangeWidthColl As Collection
    Dim oPlatePart As Object
    
    Dim oMfgPartParent As IJMfgChild
    Set oMfgPartParent = oMfgPart
    Set oPlatePart = oMfgPartParent.GetParent
    
    Dim oPlatePartHlpr  As MfgRuleHelpers.PlatePartHlpr
    Set oPlatePartHlpr = New MfgRuleHelpers.PlatePartHlpr
    
    Set oPlatePartHlpr.object = oPlatePart
    
    Dim iCount        As Long
    Dim oPlateSystem  As IJPlateSystem
    Set oPlateSystem = oPlatePartHlpr.GetRootSystem
    
    Dim bIsCanPlate As Boolean
    Dim bCylindricalPlate As Boolean
    
    bIsCanPlate = CheckIfCanPlate(oPlateSystem)
    bCylindricalPlate = IsCylidrical(oPlatePart)
        
    If bIsCanPlate = True Or bCylindricalPlate = True Then
        
        GetEligiblePortCollectionForCan oPlatePart, oPortColl
        
        If Not oPortColl Is Nothing Then
            For iCount = 1 To oPortColl.Count
                oPlatePartHlpr.UpdateConstantMargin oPortColl.Item(iCount), 7, 0.01        ' 7 - ClampMargin
            Next iCount
            Exit Sub
        End If
        
    End If
    
    'Get the eligible port collection
    GetEligiblePortAndFlangeWidthCollection oPlatePart, oPortColl, oFlangeWidthColl
    
    If oPortColl.Count = 0 Then
        Exit Sub
    End If
    
    For iCount = 1 To oPortColl.Count
        oPlatePartHlpr.UpdateConstantMargin oPortColl.Item(iCount), 7, 0.2   ' 7 - ClampMargin, 0.2 - Margin value
    Next
    
CleanUp:
    Set oMfgPartParent = Nothing
    Set oPlatePart = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1061, , "RULES")
    GoTo CleanUp
End Sub
Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)

End Sub
