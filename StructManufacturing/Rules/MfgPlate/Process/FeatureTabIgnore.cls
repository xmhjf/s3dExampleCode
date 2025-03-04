VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeatureTabIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: AntiMoldedSide
'
' Description:  Determines the process settings for the mfg plate. Only process user defined tabs while manufacturing the plate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgPlateProcess.FeatureTabIgnore"

Implements IJDMfgProcessRule
Implements IJDMfgPostUnfoldRule

Private Sub IJDMfgPostUnfoldRule_ExecuteRule(ByVal Part As Object, ByVal UnfoldedGeomColl As Object, ByVal oTabCandidateColl As IJElements, NewGeomColl As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "FeatureTabIgnore: IJDMfgPostUnfoldRule_ExecuteRule"
    On Error GoTo ErrorHandler

    'apply user created knuckle tabs
    Dim oTabHelper As IJMfgTabSymbolHelper
    Set oTabHelper = New TabSymbolHelper
 
    Dim oMfgUpdateService As MfgUpdateService
    Set oMfgUpdateService = New MfgUpdateService

    Dim oMfgTabColl As IJDObjectCollection
    Set oMfgTabColl = oMfgUpdateService.GetMfgTabs(Part)

    Dim vMoniker As Variant

    Dim oActConnection As IJDAccessMiddle
    Set oActConnection = GetActiveConnection
    
    Dim oSystemMarkFactory As New GSCADMfgSystemMark.MfgSystemMarkFactory

    Dim oPOM As IJDPOM
    Set oPOM = oActConnection.GetResourceManagerFromType("Model")

    'set relation between tab and mfgplate
    Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Dim oMfgPlatePart As IJMfgPlatePart
    Dim ePlateUpside As enumPlateSide
    
    Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
    Set oMfgPlateWrapper.object = Part
    If oMfgPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
        ePlateUpside = oMfgPlatePart.UpSide
    End If
    
    For Each vMoniker In oMfgTabColl
        Dim oTabobj As IJStrMfgTab
        Set oTabobj = oPOM.GetObject(vMoniker)
        
        Dim oPart As IJDPart
        Set oPart = oTabobj.GetCatalogPart
        
        Dim oIJDAttr As IJDAttribute
        'check if tabtype is cornertab type
        'TabType = 2
        Set oIJDAttr = GetAttribute(oPart, "IJUASMPlateTabType", "TabType")
        If oIJDAttr.Value <> 2 Then GoTo NextTab 'not a Feature skip this
             
        Dim bIsMFGPlateControlled As Boolean
        bIsMFGPlateControlled = IsTabControlledWithMfgPlate(oTabobj)
        
        If bIsMFGPlateControlled = True Then 'Rule based Feature Tab, delete it
            'delete this tab
            Dim oIJDObject As IJDObject
            Set oIJDObject = oTabobj
            oIJDObject.Remove
            'Move to next tab
            GoTo NextTab
        End If

        oTabHelper.EvaluateTabGeometriesFromTabBO Part, oTabobj, UnfoldedGeomColl, NewGeomColl
    
        'get tab moniker
        Dim oMathGeomUtilWrapper As IJDMfgGeomUtilWrapper
        Set oMathGeomUtilWrapper = New GSCADMathGeom.MfgGeomUtilWrapper
        
        Dim oTabMoniker As IMoniker
        Set oTabMoniker = oMathGeomUtilWrapper.GetMoniker(oTabobj)

        Dim lNewGeomItem As Long
        
        Dim oNamedItem As IJNamedItem
        Set oNamedItem = oTabobj

        For lNewGeomItem = 1 To NewGeomColl.GetCount
            Dim oNewGeom2d  As IJMfgGeom2d
            Set oNewGeom2d = NewGeomColl.GetGeometry(lNewGeomItem)
            
            If oNewGeom2d.GetGeometryType = STRMFG_KNUCKLE_TAB_MARK Or oNewGeom2d.GetGeometryType = STRMFG_FEATURETAB_MARK _
                Or oNewGeom2d.GetGeometryType = STRMFG_CORNERTAB_MARK Or oNewGeom2d.GetGeometryType = STRMFG_TAB_REFERENCE_MARK Then
                Dim oGeom2dMoniker As IMoniker
                Set oGeom2dMoniker = oNewGeom2d.GetMoniker
                
                
                If Not oGeom2dMoniker Is Nothing Then
                
                    Dim oMoniker1 As JDMoniker
                    Dim oMoniker2 As JDMoniker
                    Set oMoniker1 = New JDMoniker
                    Set oMoniker2 = New JDMoniker
                    
                    oMoniker1.JMoniker = oTabMoniker
                    oMoniker2.JMoniker = oGeom2dMoniker
                    
                    If oMoniker1.IsEqual(oMoniker2) Then
                        'set name for marking line
                        Dim oSystemMark As IJMfgSystemMark
                        Dim oMarkingInfo As MarkingInfo
                        
                        Set oSystemMark = oSystemMarkFactory.Create(oPOM)
                        
                        oSystemMark.Set2dGeometry oNewGeom2d
                        oSystemMark.SetMarkingSide ePlateUpside
                        
                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark
                        oMarkingInfo.Name = oNamedItem.Name '& " - " & i '"Mfg Tab - " & i
                        
                        Set oSystemMark = Nothing
                        Set oMarkingInfo = Nothing
                    End If
                    Set oMoniker1 = Nothing
                    Set oMoniker2 = Nothing
                End If
                Set oGeom2dMoniker = Nothing
            End If
        Next
        Set oMathGeomUtilWrapper = Nothing
        Set oTabMoniker = Nothing
NextTab:
    Next
    Set oSystemMarkFactory = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1085, , "RULES")
End Sub

Private Sub IJDMfgProcessRule_ExecuteRule(ByVal Part As Object, MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)

End Sub
 
