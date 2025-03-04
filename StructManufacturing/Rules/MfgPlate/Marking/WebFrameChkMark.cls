VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebFrameChkMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Web Frame Check Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.WebFrameChkMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "WebFrameChkMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_FRAMES_CHECKLINES_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
        
    ' Check if the upside is undefined
    If UpSide = UnDefinedSide Then
        UpSide = BaseSide
    End If
    
    Dim oPlatePart As IJPlatePart
    Set oPlatePart = Part
        
    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)
    
    Dim oCS As IJComplexString
    
    On Error Resume Next ' Web Frame Check Mark is not always possible to create. It depends on plate shape.
    Set oCS = CreateWebFrameCheckMarkFromParent(oPlatePart, UpSide, oConObjsCol)
     'If oCS is nothing means the previous method have failed so put it in TDL
    If oCS Is Nothing Then Exit Function
   
    On Error GoTo ErrorHandler
    Dim oGeom3d As IJMfgGeom3D
    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    'Create a SystemMark object to store additional information
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo
    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    'Set the marking side
    oSystemMark.SetMarkingSide UpSide
    
    'QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark

    oMarkingInfo.Name = "WEB FRAME CHECK MARK"

    oGeom3d.PutGeometry oCS
    oGeom3d.PutGeometrytype STRMFG_FRAMES_CHECKLINES_MARK

    'ApplyDirection oSystemMark, oConnectionData.ConnectingPort
            
    oSystemMark.Set3dGeometry oGeom3d
    
    oGeomCol3d.AddGeometry 1, oGeom3d

    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oPlatePart = Nothing
    Set oGeom3d = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oCS = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1033, , "RULES")
    GoTo CleanUp
End Function
 
