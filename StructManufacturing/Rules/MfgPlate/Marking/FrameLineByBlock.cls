VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameLineByBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2010 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Frame Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'  History:
'           09/24/2010   Suma Mallena      Created
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.FrameLineByBlock"

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    PlMrkHelpers.Initialize
    Exit Sub

ErrorHandler:
Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1010, , "RULES")
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "FrameLineByBlock: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
        
    On Error Resume Next
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(oResourceManager)
    
    CreateAPSMarkings STRMFG_FRAMELINE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    Dim dDotP           As Double
    Dim oPlateNormal    As IJDVector
    Dim oXAxis          As New DVector
    
    oXAxis.Set 1#, 0#, 0#
    Set oPlateNormal = GetPlateNeutralSurfaceNormal(Part)
    
    If Not oPlateNormal Is Nothing Then
        dDotP = oPlateNormal.Dot(oXAxis)
        If Abs(1# - Abs(dDotP)) < 0.0001 Then
            Exit Function
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    'Initialize the PlateWrapper with the PlatePart
    Dim oPlate As StructDetailObjects.PlatePart
    Set oPlate = New StructDetailObjects.PlatePart
    Set oPlate.object = Part
    
    'Get the Plate Part Surface
    Dim oOffsetPort             As IJPort
    Dim oPlateSurfGeom          As Object
    
    Set oOffsetPort = oPlate.BasePort(BPT_Offset)
    Set oPlateSurfGeom = oOffsetPort.Geometry
    
    Dim nIndex                  As Long
    Dim oConObjsCol             As Collection
    Dim oRefplanes              As Collection

    ' LEAD_BLOCK_REFERENCE  -- Gets the frames only from the immediate parent block
    m_oMfgRuleHelper.GetBlockFrames Part, oPlateSurfGeom, x, LEAD_BLOCK_REFERENCE, oRefplanes, oConObjsCol

    If Not (oConObjsCol Is Nothing) Then
    
        Dim oSPGMiddleHelper    As New SP3dGSMiddleHelper.SPGMiddleHelper
        Dim oFrameSystem        As IJDCoordinateSystem
        Dim oFrameSysMoniker    As IMoniker

        If oConObjsCol.Count > 0 Then
            'Get the Frame System
            oSPGMiddleHelper.GetCoordinateSystemOfPlane oRefplanes(1), oFrameSystem
            Set oFrameSysMoniker = m_oMfgRuleHelper.GetMoniker(oFrameSystem)
        End If

        On Error Resume Next
        For nIndex = 1 To oConObjsCol.Count
            
            Dim oFrameline          As IUnknown
            Dim oCurves             As IJElements
            Dim oCS                 As IJComplexString
            Dim oSystemMark         As IJMfgSystemMark
            Dim oMarkingInfo        As MarkingInfo
            Dim oNamedItem          As IJNamedItem
            Dim oGeom3d             As IJMfgGeom3D
            
            Set oFrameline = oConObjsCol.Item(nIndex)
            
            ' Get type of Frame and set as custom attribute on Marking Info
            Dim oGridData As ISPGGridData
            Set oGridData = oRefplanes.Item(nIndex)
            
            Dim oCodeListMetaData As IJDCodeListMetaData
            Set oCodeListMetaData = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
            Dim strPlaneType As String
            strPlaneType = oCodeListMetaData.LongStringValue("GridPlaneType", oGridData.Type)
            
    '       Convert the IJWireBody to a collection of IJComplexString's
            Set oCurves = m_oMfgRuleHelper.WireBodyToComplexStrings(oFrameline)

            For Each oCS In oCurves

    '           Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

    '           Set the marking side
                oSystemMark.SetMarkingSide UpSide

    '           QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark

    '           Get the Frame Name
                Set oNamedItem = oRefplanes(nIndex)
                oMarkingInfo.Name = oNamedItem.Name
                
                'Set the Frame Type
                If Not strPlaneType = vbNullString Then
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", strPlaneType
                End If

                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_FRAMELINE_MARK
                
    '           Set the Frame System Moniker on the Geom3d object
                oGeom3d.PutMoniker oFrameSysMoniker
                oGeom3d.Put3dModelMoniker m_oMfgRuleHelper.GetMoniker(oRefplanes.Item(nIndex))

                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry 1, oGeom3d

                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oGeom3d = Nothing
                Set oNamedItem = Nothing
            Next oCS
            
            Set oGridData = Nothing
            Set oCodeListMetaData = Nothing
        Next nIndex
        
        Set oFrameSystem = Nothing
        Set oFrameSysMoniker = Nothing
    End If

'    Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:
    
    Set oOffsetPort = Nothing
    Set oPlateSurfGeom = Nothing
    Set oCS = Nothing
    Set oFrameline = Nothing
    Set oConObjsCol = Nothing
    Set oSystemMark = Nothing
    Set oGeom3d = Nothing
    Set oResourceManager = Nothing
    Set oSPGMiddleHelper = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1010, , "RULES")
    GoTo CleanUp
End Function
