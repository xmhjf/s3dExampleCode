VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.SeamControl"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "SeamControl: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler

    Dim SeamCurves As IJElements
    If TypeOf Part Is ISPSMemberPartPrismatic Then
        Dim oMemberPart As ISPSMemberPartPrismatic
        Set oMemberPart = Part
        
        Dim MemberSeamInfo As ISPSMemberPrismaticSeamData
        Set MemberSeamInfo = oMemberPart.SeamData
        
            If Not MemberSeamInfo Is Nothing Then
                Set SeamCurves = MemberSeamInfo.SeamDisplayCurves
            End If
    End If

    If SeamCurves Is Nothing Then GoTo CleanUp
    If SeamCurves.Count = 0 Then GoTo CleanUp
    
    Dim oMfgGeomHelper As New MfgGeomHelper
    
    Dim oCScollection As IJElements
    Set oCScollection = oMfgGeomHelper.OptimizedMergingOfInputCurves(SeamCurves)
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
        
    Dim oCS As IJComplexString
    For Each oCS In oCScollection
        Dim oGeom3d As IJMfgGeom3d
        Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
        oGeom3d.PutGeometry oCS
        oGeom3d.PutGeometrytype STRMFG_COMMON_SEAM_MARK
        oGeom3d.FaceId = UpSide
    
        'Create a SystemMark object to store additional information
        Dim oSystemMark As IJMfgSystemMark
        Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
        'Set the marking side
        oSystemMark.SetMarkingSide UpSide
    
        'QI for the MarkingInfo object on the SystemMark
        Dim oMarkingInfo As MarkingInfo
        Set oMarkingInfo = oSystemMark
    
        oMarkingInfo.name = "MEMBER TUBE SEAM"
    
        oSystemMark.Set3dGeometry oGeom3d
    
        oGeomCol3d.AddGeometry 1, oGeom3d
    Next
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oMemberPart = Nothing
    Set MemberSeamInfo = Nothing
    Set SeamCurves = Nothing
    Set oResourceManager = Nothing
    Set oGeom3d = Nothing
    Set oGeomCol3d = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2015, , "RULES")
    GoTo CleanUp
End Function

 
