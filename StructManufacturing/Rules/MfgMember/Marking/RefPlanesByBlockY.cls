VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefPlanesByBlockY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg member
'
'  History:
'           09/24/2010   Suma Mallena      Created
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.RefPlanesByBlockY"

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
    Const METHOD = "RefPlanesByBlockY: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oSDMemberWrapper As Object
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oMfgProfileOutput As IJMfgProfileOutput
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    If Not oProfileWrapper.ProfileHasMfgPart(oMfgProfilePart) Then
        Exit Function
    End If
    Set oMfgProfileOutput = oMfgProfilePart
    Dim oFaces() As Variant
    Dim varFace As Variant
    Dim lFace As Long
    oMfgProfileOutput.GetPartFaces StructMfgMarkingSide, StructMfgAsAny, oFaces
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    If UBound(oFaces) >= LBound(oFaces) Then
        For Each varFace In oFaces
            lFace = varFace
            CreateFrameMarksByBlock oGeomCol3d, lFace, "Y", oSDMemberWrapper, ReferenceObjColl
        Next varFace
    Else
        CreateFrameMarksByBlock oGeomCol3d, UpSide, "Y", oSDMemberWrapper, ReferenceObjColl
    End If
'    Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:
    Set oSDMemberWrapper = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2012, , "RULES")
    GoTo CleanUp
End Function

