VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateEdgeCheckMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Plate Edge check Markings in the MfgPlate command.
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

Private Const MODULE = "MfgPlateMarking.PlateEdgeCheckMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "PlateEdgeCheckMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_EDGE_CHECKLINES_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)
    
    If oConObjsCol Is Nothing Then
        'No connected objects so we can leave
        GoTo CleanUp
    ElseIf oConObjsCol.Count = 0 Then
        'No connected objects so we can leave
        GoTo CleanUp
    End If
    
    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part
    
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart

    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    Dim oPlateCreation_AE As IJMfgPlateCreation_AE
    Set oPlateCreation_AE = oMfgPart.ActiveEntity
    
    Dim oOuteContouGeomCol3d As IJMfgGeomCol3d
    Set oOuteContouGeomCol3d = oPlateCreation_AE.GeometriesBeforeUnfold
    If oOuteContouGeomCol3d Is Nothing Then
        'Since we do not find any geometries to be marked we can exit
        Exit Function
    ElseIf oOuteContouGeomCol3d.GetCount = 0 Then
        'zero elements count, so we can leave
        GoTo CleanUp
    End If
        
    Dim oPlateSys As IJPlate
    Set oPlateSys = Part
    If Not oPlateSys.plateType = TBulkheadPlate Then Exit Function

    Dim oPositionCol As Collection
    Set oPositionCol = oPlateWrapper.GetConnectingHullSeamPositions(oOuteContouGeomCol3d, oConObjsCol)
    
    If (oPositionCol Is Nothing) Then
        'Since there are no seams to be taken into account we can leave the function
        GoTo CleanUp
    End If
    
    If (oPositionCol.Count = 0) Then
        'Since there are no seams to be taken into account we can leave the function
        GoTo CleanUp
    End If
    
    Dim i As Long
    
    For i = 1 To oPositionCol.Count
        Dim oSystemMark As IJMfgSystemMark
        Dim oObjSystemMark As IUnknown
        Dim oMoniker As IMoniker
        Dim oMarkingInfo As MarkingInfo
        Dim oGeom3d As IJMfgGeom3D
        Dim oPosition As IJDPosition
        Dim oCS As IJComplexString
        
        Set oPosition = oPositionCol.Item(i)
        
        Set oCS = CreateEdgeCheckMark(Part, oPosition, UpSide)

        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        'Set the marking side
        oSystemMark.SetMarkingSide UpSide

        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark

        oMarkingInfo.Name = "EDGE_CHECK_MARK"

        Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oGeom3d.PutGeometry oCS
        oGeom3d.PutGeometrytype STRMFG_EDGE_CHECKLINES_MARK
        Set oObjSystemMark = oSystemMark
        
        oSystemMark.Set3dGeometry oGeom3d
                
        oGeomCol3d.AddGeometry 1, oGeom3d

NextItem:
        Set oPosition = Nothing
        Set oCS = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oObjSystemMark = Nothing
    Next i
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oOuteContouGeomCol3d = Nothing
    Set oPlateWrapper = Nothing
    Set oPositionCol = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1017, , "RULES")
    GoTo CleanUp
End Function

 
