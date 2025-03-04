VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginCalculated"
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
' The rule handles margin creation for Can plates, flange brackets
' Author:
'
' Comments:
' 2010.06.22    Siva    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const Module = "MfgPlateProcess"
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "MarginCalculated: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error Resume Next
    
    Dim oPlatePart As Object
    
    Dim oMfgPartParent As IJMfgChild
    Set oMfgPartParent = oMfgPart
    Set oPlatePart = oMfgPartParent.GetParent
    
    Dim oPlatePartHlpr  As MfgRuleHelpers.PlatePartHlpr
    Set oPlatePartHlpr = New MfgRuleHelpers.PlatePartHlpr
    
    Set oPlatePartHlpr.object = oPlatePart
    
    'First remove the existing rule based clamp margins.
    oPlatePartHlpr.RemoveMargins 7     'ClampMargin
    
    Dim oPortColl As IJElements
    Dim iCount As Long
    
    Dim oPlateSystem    As IJPlateSystem
    Set oPlateSystem = oPlatePartHlpr.GetRootSystem
    
    Dim dPlateThickness As Double
    Dim oPlate As IJPlate
    Set oPlate = oPlateSystem
    
    dPlateThickness = oPlate.Thickness
    
    'Check if the root plate system is Can plate
    
    Dim bIsCanPlate As Boolean
    Dim bCylindricalPlate As Boolean
    
    bIsCanPlate = CheckIfCanPlate(oPlateSystem)
    bCylindricalPlate = IsCylidrical(oPlatePart)
        
    If bIsCanPlate = True Or bCylindricalPlate = True Then
        
        GetEligiblePortCollectionForCan oPlatePart, oPortColl
        
        If Not oPortColl Is Nothing Then
            
           If dPlateThickness >= 0.03 Then
             For iCount = 1 To oPortColl.Count
                 oPlatePartHlpr.UpdateConstantMargin oPortColl.Item(iCount), 7, 0.01        ' 7 - ClampMargin
             Next iCount
           End If

           Exit Sub

        End If
        
    End If
    
    Dim oFlangeWidthColl As Collection
    
    ' Get the eligible port collection and flange width collection
    GetEligiblePortAndFlangeWidthCollection oPlatePart, oPortColl, oFlangeWidthColl
    
    If oPortColl.Count = 0 Then
        Exit Sub
    End If
    
    For iCount = 1 To oPortColl.Count
        
        Dim dFlangeWidth As Double, dMargin As Double
        
        dFlangeWidth = oFlangeWidthColl.Item(iCount)
                      
        dMargin = 0#
        If dPlateThickness <= 0.007 Then
            If dFlangeWidth < 0.05 Then
                dMargin = 0.075 - dFlangeWidth
            End If
        Else
            If dFlangeWidth < 0.075 Then
                dMargin = 0.075 - dFlangeWidth
            End If
        
        End If
        
        If dMargin <> 0# Then
            oPlatePartHlpr.UpdateConstantMargin oPortColl.Item(iCount), 7, dMargin       ' 7 - ClampMargin
        End If
    Next
    

CleanUp:
    Set oMfgPartParent = Nothing
    Set oPlatePart = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1060, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)

End Sub
