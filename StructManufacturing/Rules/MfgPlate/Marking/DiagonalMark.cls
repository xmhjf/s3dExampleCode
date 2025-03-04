VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DiagonalMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Diagonal Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'       TBH     feb. 6. 2002    created
'       KONI    june 25 2002    Status - work when pinjig is created by platepart (not assembly)
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.DiagonalMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "DiagonalMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_DIAGONALS_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    'Check if part has pin jig related to it.
    If oPlateWrapper.HasMfgPinJig = False Then Exit Function
    
    'Get the diagonal lines from the PinJig object
    Dim oDiagonalLines As IJMfgGeomCol3d
    Set oDiagonalLines = oPlateWrapper.DiagonalLines
    
    If oDiagonalLines Is Nothing Then
        'Since there are no diagonal lines we can leave the function
        GoTo CleanUp
    End If
    
    Dim nIndex As Long
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oGeom3dLine As IJMfgGeom3D
    Dim oSystemMark As IJMfgSystemMark
    
    'Can't use "For Each" when user defined types in collection
    For nIndex = 1 To oDiagonalLines.GetCount
        Set oGeom3dLine = oDiagonalLines.GetGeometry(1)

        oDiagonalLines.RemoveGeometry oGeom3dLine

        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        'Set the marking side
        oSystemMark.SetMarkingSide UpSide

        oGeom3dLine.PutGeometrytype STRMFG_DIAGONALS_MARK

        oSystemMark.Set3dGeometry oGeom3dLine
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' To Do: Since the returned DiagonalLines are not persistent objects,
        '        we need to make them as persistent objects somewhere in this function.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        oGeomCol3d.AddGeometry 1, oGeom3dLine

        Set oObjSystemMark = Nothing
        Set oMoniker = Nothing
        Set oGeom3dLine = Nothing
        Set oSystemMark = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oPlateWrapper = Nothing
    Set oDiagonalLines = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1008, , "RULES")
    GoTo CleanUp
End Function

 
