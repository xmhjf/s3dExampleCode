VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InverseBendingLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.InverseBendingLine"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    
    Const METHOD = "InverseBendingLine: IJDMfgSystemMarkingRule_CreateAfterUnfold"
    On Error GoTo ErrorHandler
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    Dim oMfgPart As IJMfgProfilePart
    
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    'retrieve collection of lines
    Dim oIBLColl As IJMfgGeomCol2d
    Set oIBLColl = oMfgProfileWrapper.GetBendingLines
    
    If oIBLColl Is Nothing Then
        'Since there are no IBL we can leave the function
        GoTo CleanUp
    End If

    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    'Create the 2d collection
    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = m_oGeomCol2dFactory.Create(oResourceManager)
    
    Dim iIndex As Integer, iGeomCount As Integer
    iIndex = 1
    iGeomCount = oIBLColl.Getcount
    'Iterate through the collection of lines
    While iIndex <= iGeomCount
        'check for the type of geom (STRMFG_BENDING_LINE or STRMFG_TOPFLANGE_BENDING_LINE or STRMFG_BOTTOMFLANGE_BENDING_LINE)
        'with Face Ids it should suffice to check for the type of geom STRMFG_BENDING_LINE
        Dim oMfgGeom2d As IJMfgGeom2d
        Set oMfgGeom2d = oIBLColl.GetGeometry(iIndex)
        
        Dim eGeometryType As StrMfgGeometryType
        eGeometryType = oMfgGeom2d.GetGeometryType
        
        Dim lFaceId As Long
        lFaceId = oMfgGeom2d.FaceId
        
        If (eGeometryType = STRMFG_BENDING_LINE) Then
            'Can only be in one collection at the time
            oIBLColl.RemoveGeometry oMfgGeom2d
            iIndex = iIndex - 1
            iGeomCount = oIBLColl.Getcount
            
            Dim oSystemMark As IJMfgSystemMark
            ' Check if unfold algorithm created a system mark already!
            Set oSystemMark = oMfgGeom2d.SystemMark
            If oSystemMark Is Nothing Then
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                oSystemMark.SetMarkingSide lFaceId
                oSystemMark.Set2dGeometry oMfgGeom2d
            ElseIf oSystemMark.GetMarkingSide <> lFaceId Then
                LogMessage Err, MODULE, METHOD, "Internally created system mark has wrong marking side for IBL!"
                oSystemMark.SetMarkingSide lFaceId
            End If
        
            'QI for the MarkingInfo object on the SystemMark
            Dim oMarkingInfo As IJMarkingInfo
            Set oMarkingInfo = oSystemMark
            
            Dim IBLname As String
            IBLname = "BENDING LINE"
        
            Dim IBLidx As Variant
            oMarkingInfo.GetAttributeValue "IBL_INDEX", IBLidx
            If Not IsEmpty(IBLidx) Then
                IBLname = "IBL-" & CStr(IBLidx)
            End If

            'Set marking information, if any
            oMarkingInfo.name = IBLname
            
            Dim oObjSystemMark As IUnknown
            Dim oMoniker As IMoniker
           
            Set oObjSystemMark = oSystemMark
    
            oMfgGeom2d.PutGeometrytype eGeometryType
    
            oGeomCol2d.AddGeometry 1, oMfgGeom2d
        ElseIf (eGeometryType = STRMFG_BENDING_CONTROLLINES_MARK) Then
            'Can only be in one collection at the time
            oIBLColl.RemoveGeometry oMfgGeom2d
            iIndex = iIndex - 1
            iGeomCount = oIBLColl.Getcount
            oGeomCol2d.AddGeometry 1, oMfgGeom2d
        End If
    
        Set oMfgGeom2d = Nothing
        Set oSystemMark = Nothing
        Set oObjSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oMoniker = Nothing
        iIndex = iIndex + 1
    Wend
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateAfterUnfold = oGeomCol2d
    
CleanUp:
    If Not oIBLColl Is Nothing Then
        Dim oObject As IJDObject
        Set oObject = oIBLColl
        oObject.Remove
        Set oObject = Nothing
    End If
    
    Set oMfgProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oMfgPart = Nothing
    Set oIBLColl = Nothing


    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2005, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d

End Function

 
