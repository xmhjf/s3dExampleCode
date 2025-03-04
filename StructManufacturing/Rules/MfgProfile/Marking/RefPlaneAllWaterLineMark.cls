VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefPlanesAllWL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
'
'  History:
'           09/24/2010   Suma Mallena      Created
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.RefPlanesAllWL"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "RefPlanesAllWL: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_WATERLINE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

    'Initialize the ProfileWrapper
    Dim oSDProfileWrapper As Object
    'Create the SD profile Wrapper and initialize it
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = Part

    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oCSColl         As IJElements
    Dim oMfgGeomHelper  As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
     
    Dim oMfgFrameSystem As IJDCoordinateSystem
    Dim nIndex          As Long
    
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
        End If
    Next nIndex
    
    ' Within V7, the Mfg object inherit the Frame System. When not found we do
    ' leave by the default rule. Customers can activate the Main CS.
    ' TR-CP·121466  manufacturing Plate profile being marked even though Frame system not defined
    If oMfgFrameSystem Is Nothing Then
        Exit Function
    End If
    
    Dim oFrameSysMoniker As IMoniker
    Set oFrameSysMoniker = m_oMfgRuleHelper.GetMoniker(oMfgFrameSystem)
    
    Dim oMfgHFrameSys As IHFrameSystem
    Set oMfgHFrameSys = oMfgFrameSystem
    
    'Get the frame line intersections using the assumption that there is a coordinate system called "Main CS"
    'This is mandatory for all the Hicadec unwrapping algorithms therefore this seems a safe bet.
    'Since it is WaterLines it automatically means the "Z" direction

    Dim oConObjsCol     As Collection
    Dim oRefplanes      As Collection
    
    oSDProfileWrapper.ReferencePlaneRange oMfgHFrameSys.name, "Z", oRefplanes, oConObjsCol, JXSEC_WEB_LEFT

    If Not (oConObjsCol Is Nothing) Then
        
        On Error Resume Next
        For nIndex = 1 To oConObjsCol.Count
        
            Dim oWB             As IJWireBody
            Dim oCS             As IJComplexString
            Dim oSystemMark     As IJMfgSystemMark
            Dim oMarkingInfo    As MarkingInfo
            Dim oGeom3d         As IJMfgGeom3d
    
            Set oWB = oConObjsCol.Item(nIndex)
    
        '   Convert the IJWireBody to a IJComplexString
            Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWB)
            
            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If
                
            If (oCSColl Is Nothing) Then
                GoTo NextItem
            End If
            
            ' Get type of Frame and set as custom attribute on Marking Info
            Dim oGridData As ISPGGridData
            Set oGridData = oMfgHFrameSys.PrincipalZAxis.Frame(oRefplanes.Item(nIndex))
       
            Dim oCodeListMetaData As IJDCodeListMetaData
            Set oCodeListMetaData = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

            Dim strPlaneType As String
            strPlaneType = oCodeListMetaData.LongStringValue("ElevPlaneType", oGridData.Type)
            
            For Each oCS In oCSColl
            '   Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    
            '   Set the marking side
                oSystemMark.SetMarkingSide UpSide
    
            '   QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
    
            '   Get the Frame Name
                oMarkingInfo.name = oRefplanes(nIndex)
                
                'Set the Frame Type
                If Not strPlaneType = vbNullString Then
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", strPlaneType
                End If
    
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_WATERLINE_MARK
                oGeom3d.FaceId = UpSide

                oSystemMark.Set3dGeometry oGeom3d
                
                oGeom3d.PutMoniker oFrameSysMoniker
                oGeom3d.Put3dModelMoniker m_oMfgRuleHelper.GetMoniker(oMfgHFrameSys.PrincipalZAxis.Frame(oRefplanes.Item(nIndex)))
                oGeomCol3d.AddGeometry 1, oGeom3d
                
                Set oSystemMark = Nothing
                Set oGeom3d = Nothing
                Set oMarkingInfo = Nothing
                Set oCS = Nothing
                
            Next oCS
NextItem:
    
            Set oWB = Nothing
            Set oCS = Nothing
            Set oSystemMark = Nothing
            Set oGeom3d = Nothing
            Set oMarkingInfo = Nothing
            Set oGridData = Nothing
            Set oCodeListMetaData = Nothing
        Next nIndex
        
    End If

'    Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oRefplanes = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2013, , "RULES")
    GoTo CleanUp
End Function
