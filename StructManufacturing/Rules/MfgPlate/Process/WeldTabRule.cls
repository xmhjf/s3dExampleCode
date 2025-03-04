VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldTabRule"
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
' Description:  Determines the process settings for the mfg plate. Process user created tabs and create rule proposed tabs
'               Establish a relation between the rule created tab and the MfgPlate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const MODULE As String = "MfgPlateProcess.FeatureTabRule"

Private Const FEATURETAB_MARKINGNAME = "FEATURETAB_MARKING"
Implements IJDMfgProcessRule
Implements IJDMfgPostUnfoldRule

Private Sub IJDMfgPostUnfoldRule_ExecuteRule(ByVal Part As Object, ByVal UnfoldedGeomColl As Object, ByVal oTabCandidateColl As IJElements, NewGeomColl As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "FeatureTabRule: IJDMfgProcessRule_ExecuteRule"
    On Error GoTo ErrorHandler
    
    'apply user created knuckle tabs
    
    Dim oTabHelper As IJMfgTabSymbolHelper

    Dim oMfgUpdateService As MfgUpdateService
    Set oMfgUpdateService = New MfgUpdateService

    Dim oMfgTabColl As IJDObjectCollection
    Set oMfgTabColl = oMfgUpdateService.GetMfgTabs(Part)

    Dim oMathGeomUtilWrapper As IJDMfgGeomUtilWrapper
    Set oMathGeomUtilWrapper = New GSCADMathGeom.MfgGeomUtilWrapper
            
    Dim vMoniker As Variant

    Dim oActConnection As IJDAccessMiddle
    Set oActConnection = GetActiveConnection

    Dim oPOM As IJDPOM
    Set oPOM = oActConnection.GetResourceManagerFromType("Model")
    Dim oSystemMarkFactory As New GSCADMfgSystemMark.MfgSystemMarkFactory
    Dim lNewGeomItem As Long

    'set relation between tab and mfgplate
    Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Dim oMfgPlatePart As IJMfgPlatePart
    Dim ePlateUpside As enumPlateSide
    Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
    Set oMfgPlateWrapper.object = Part
    If oMfgPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
        ePlateUpside = oMfgPlatePart.UpSide
    End If

    Dim oIJDAttr As IJDAttribute

    For Each vMoniker In oMfgTabColl
        Dim oTabobj As IJStrMfgTab
        Set oTabobj = oPOM.GetObject(vMoniker)
        Dim oPart As IJDPart
        Set oPart = oTabobj.GetCatalogPart
        
        'check if tabtype is Weldtab type
        Set oIJDAttr = GetAttribute(oPart, "IJUASMPlateTabType", "TabType")
        If oIJDAttr.Value = 3 Or oIJDAttr.Value = 4 Then
            Set oTabHelper = New TabSymbolHelper
            oTabHelper.EvaluateTabGeometriesFromTabBO Part, oTabobj, UnfoldedGeomColl, NewGeomColl
            Set oTabHelper = Nothing
        End If
        
        Dim oNI As IJNamedItem
        Set oNI = oTabobj
            
        Dim oTabMoniker As IMoniker
        Set oTabMoniker = oMathGeomUtilWrapper.GetMoniker(oTabobj)
        
        For lNewGeomItem = 1 To NewGeomColl.GetCount
            Dim oNewGeom2dBO  As IJMfgGeom2d
            Set oNewGeom2dBO = NewGeomColl.GetGeometry(lNewGeomItem)
            
            If oNewGeom2dBO.GetGeometryType = STRMFG_KNUCKLE_TAB_MARK Or oNewGeom2dBO.GetGeometryType = STRMFG_FEATURETAB_MARK _
                Or oNewGeom2dBO.GetGeometryType = STRMFG_CORNERTAB_MARK Or oNewGeom2dBO.GetGeometryType = STRMFG_TAB_REFERENCE_MARK Then
                Dim oGeom2dMoniker As IMoniker
                Set oGeom2dMoniker = oNewGeom2dBO.GetMoniker
                
                If Not oGeom2dMoniker Is Nothing Then
                    Dim oMoniker1 As JDMoniker
                    Dim oMoniker2 As JDMoniker
                    Set oMoniker1 = New JDMoniker
                    Set oMoniker2 = New JDMoniker
                    
                    oMoniker1.JMoniker = oTabMoniker
                    oMoniker2.JMoniker = oGeom2dMoniker
                    
                    If oMoniker1.IsEqual(oMoniker2) Then
                    
                        'set name for marking line
                        Dim oSystemMarkBO As IJMfgSystemMark
                        Dim oMarkingInfoBO As MarkingInfo
                        
                        Set oSystemMarkBO = oSystemMarkFactory.Create(oPOM)
                        
                        oSystemMarkBO.Set2dGeometry oNewGeom2dBO
                        oSystemMarkBO.SetMarkingSide ePlateUpside
                        
                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfoBO = oSystemMarkBO
                        oMarkingInfoBO.Name = oNI.Name '& " - " & i '"Mfg Tab - " & i
                        
                        Set oSystemMarkBO = Nothing
                        Set oMarkingInfoBO = Nothing
                    End If
                    Set oMoniker1 = Nothing
                    Set oMoniker2 = Nothing
                End If
                Set oGeom2dMoniker = Nothing
            End If
        Next
        Set oTabMoniker = Nothing
        Set oNI = Nothing
        Set oIJDAttr = Nothing
    Next
    
    On Error Resume Next
    Dim oWeldTabSmartOccRoot As IJSmartClass
    Dim oWeldTabSelRule As IJMfgTabSelectionHelper
    GetSmartOccRootClassGivenSCtype 10028, 400, "SMTabWeld", oWeldTabSmartOccRoot, oWeldTabSelRule
    On Error GoTo ErrorHandler
    
    'For models migrate from v2011R1,SMTabWeld smartclass does not exist
    If oWeldTabSelRule Is Nothing Then
        GetSmartOccRootClassGivenSCtype 10028, 100, "SMTabCorners", oWeldTabSmartOccRoot, oWeldTabSelRule
    End If

    Dim i As Integer
    For i = 1 To oTabCandidateColl.Count
        Dim oTabCandidate As IJMfgTabCandidate
        Set oTabCandidate = oTabCandidateColl.Item(i)

        Dim oSOPartClass As IJDSymbolDefHelper
        Dim oTabPart As IJDPart

        Dim oCatalogQuery As IJSRDQuery
        Set oCatalogQuery = New SRDQuery

        Dim oSOFeatureClassQuery As IJSmartQuery
        Set oSOFeatureClassQuery = oCatalogQuery
        
        Dim oHelper As IJPartSupport
        Set oHelper = New PlatePartSupport
        Set oHelper.Part = Part
        
        If oTabCandidate.NumberOfKnuckles = 0 Then
            'check if any of the contours is a feature or an opening
            Dim Operator(1 To 2) As Object
            On Error Resume Next
            oHelper.GetPortOperator GetPlatePortGivenMfgContour(oTabCandidate.FirstContour), Operator(1)
            oHelper.GetPortOperator GetPlatePortGivenMfgContour(oTabCandidate.SecondContour), Operator(2)

            If TypeOf Operator(1) Is IJSeam Or _
                TypeOf Operator(2) Is IJSeam Then
                    
                Set oSOPartClass = oWeldTabSmartOccRoot
                Set oTabPart = oWeldTabSelRule.SelectTabCatalogPart(oSOPartClass, oTabCandidate, 3, UnfoldedGeomColl)
                                    
                If TypeOf Operator(1) Is IJSeam And _
                    TypeOf Operator(2) Is IJSeam Then
                    
                    Dim Length(1 To 2) As Double
                    Length(1) = GetLengthOfGeom2dSegment(oTabCandidate.FirstContour)
                    Length(2) = GetLengthOfGeom2dSegment(oTabCandidate.SecondContour)
                
                    If Length(2) < Length(1) Then
                       
                       oTabCandidate.IsFirstContourDriver = False
                    Else
                        oTabCandidate.IsFirstContourDriver = True
                    End If
                Else
                    If TypeOf Operator(1) Is IJSeam Then
                       oTabCandidate.IsFirstContourDriver = True
                    Else
                        oTabCandidate.IsFirstContourDriver = False
                    End If
                
                End If
                                
            End If
            Set Operator(2) = Nothing
            Set Operator(1) = Nothing
            On Error GoTo ErrorHandler
        End If
        
        If Not oTabPart Is Nothing Then
        
            'check if tabtype is Weldtab type
            Set oIJDAttr = GetAttribute(oTabPart, "IJUASMPlateTabType", "TabType")
            If oIJDAttr.Value = 3 Then
        
                ' New object to be created each iteration
                Dim oMfgTab As IJStrMfgTab
                
                Set oTabHelper = New TabSymbolHelper
                oTabHelper.EvaluateTabGeometriesFromTabCandidates Part, oSOPartClass, oTabPart, oTabCandidate, UnfoldedGeomColl, NewGeomColl, oMfgTab
                Set oTabHelper = Nothing
                
                If oMfgTab Is Nothing Then
                    Set oTabPart = Nothing
                    GoTo NextItem
                Else
                    'assign a name to the tab created
                    Dim oNamedItem As IJNamedItem
                    Set oNamedItem = oMfgTab
                    oNamedItem.Name = oTabPart.PartNumber & " - " & i 'oNamedItem.Name = "Mfg Tab - " & i

                    Dim oTabBevel As IJMfgTabBevel
                    Set oTabBevel = oMfgTab
                    If Not oTabBevel Is Nothing Then
                        oTabBevel.BevelOnMarking = 0 ' Ignore
                        oTabBevel.BevelOnContour = 7 ' Apply on Driving and Driven contours
                        Set oTabBevel = Nothing
                    End If
                    
                    'get tab moniker
                    Dim oTabBOMoniker As IMoniker
                    Set oTabBOMoniker = oMathGeomUtilWrapper.GetMoniker(oMfgTab)

                    For lNewGeomItem = 1 To NewGeomColl.GetCount
                        Dim oNewGeom2d  As IJMfgGeom2d
                        Set oNewGeom2d = NewGeomColl.GetGeometry(lNewGeomItem)
                        
                        If oNewGeom2d.GetGeometryType = STRMFG_KNUCKLE_TAB_MARK Or oNewGeom2d.GetGeometryType = STRMFG_FEATURETAB_MARK _
                            Or oNewGeom2d.GetGeometryType = STRMFG_CORNERTAB_MARK Or oNewGeom2d.GetGeometryType = STRMFG_TAB_REFERENCE_MARK Then
                            
                            Dim oGeom2dMarkMoniker As IMoniker
                            Set oGeom2dMarkMoniker = oNewGeom2d.GetMoniker
                    
                            If Not oGeom2dMarkMoniker Is Nothing Then
                                Dim oMonikerTab As JDMoniker
                                Dim oMonikerGeom2d As JDMoniker
                                Set oMonikerTab = New JDMoniker
                                Set oMonikerGeom2d = New JDMoniker
                                
                                oMonikerTab.JMoniker = oTabBOMoniker
                                oMonikerGeom2d.JMoniker = oGeom2dMarkMoniker
                                
                                If oMonikerTab.IsEqual(oMonikerGeom2d) Then
                
                                    'set name for marking line
                                    Dim oSystemMark As IJMfgSystemMark
                                    Dim oMarkingInfo As MarkingInfo
                                    
                                    Set oSystemMark = oSystemMarkFactory.Create(oPOM)
                                    
                                    oSystemMark.Set2dGeometry oNewGeom2d
                                    oSystemMark.SetMarkingSide ePlateUpside
                                    
                                    'QI for the MarkingInfo object on the SystemMark
                                    Set oMarkingInfo = oSystemMark
                                    oMarkingInfo.Name = oTabPart.PartNumber & " - " & i '"Mfg Tab - " & i
                                    
                                    Set oSystemMark = Nothing
                                    Set oMarkingInfo = Nothing
                                End If
                                Set oMonikerTab = Nothing
                                Set oMonikerGeom2d = Nothing
                            End If
                            Set oGeom2dMarkMoniker = Nothing
                        End If
                    Next
                    Set oTabBOMoniker = Nothing
                End If
                
                If Not oMfgPlatePart Is Nothing Then
                    oMfgPlatePart.AddMfgDefinition oMfgTab
                End If
    
                Set oTabPart = Nothing
            End If
            Set oIJDAttr = Nothing
        
        End If
        Set oSOPartClass = Nothing
        Set oTabCandidate = Nothing
NextItem:
    Next
    
CleanUp:
    Set oTabCandidate = Nothing
    Set oHelper = Nothing
    Set oMathGeomUtilWrapper = Nothing
    
    Set oWeldTabSmartOccRoot = Nothing
    Set oWeldTabSelRule = Nothing
    Set oSystemMarkFactory = Nothing
    Set oMfgPlatePart = Nothing
    Set oMfgPlateWrapper = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1092, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJDMfgProcessRule_ExecuteRule(ByVal Part As Object, MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    
End Sub

