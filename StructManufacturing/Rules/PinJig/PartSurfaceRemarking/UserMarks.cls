VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigPartSurfLCRule
' Module: UserMarks
'
' Description: The lines added via marking command that is meant to show up as remarking
'              lines AS-IS (without alteration) in the pinjig will be sent to the semantic.
'
'*******************************************************************************

Option Explicit
Private Const MODULE As String = "MfgPinJigPartSurfLCRule.UserMarks"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements
    Const METHOD As String = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim oPinJig As IJPinJig
    ' Note that, when execution gets here, this input Pin jig is still only partially computed!
    Set oPinJig = PartialPinJig
    
    ' If set to True, the below routine will get only those user marks placed on input pin jig
    ' If set to False, the below routine will get all user marks on pin jig's supported plates
    Const OnlyThisPinJig As Boolean = False
    
    Dim oUserMarkOnThisPinJig As IJElements
    ' Return only user marks of type "STRMFG_PINJIG_MARK_DIAGONAL"
    Set oUserMarkOnThisPinJig = oPinJig.MarkingLinesOnSupportedPlates(OnlyThisPinJig, STRMFG_PINJIG_DIAGONAL, PinJigRemarkingSide)
    
    Dim oColl2 As IJElements
    ' Return only user marks of type "STRMFG_PINJIG_MARK"
    Set oColl2 = oPinJig.MarkingLinesOnSupportedPlates(OnlyThisPinJig, STRMFG_PINJIG_MARK, PinJigRemarkingSide)
    Set oPinJig = Nothing
    
    ' Append collection with STRMFG_PINJIG_MARK marks to User mark collection with STRMFG_PINJIG_MARK_DIAGONAL marks
    oUserMarkOnThisPinJig.AddElements oColl2
    Set oColl2 = Nothing
    
    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = oUserMarkOnThisPinJig
        
    Set oUserMarkOnThisPinJig = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5008, , "RULES")

End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsUserRemark]"
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD As String = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject
    
    Dim oUserMarkColl As IJElements
    Set oUserMarkColl = New JObjectCollection
    
    Dim i As Long
    For i = 1 To ElemsToRemark.Count
        If TypeOf ElemsToRemark.Item(i) Is IJMfgMarkingLines_AE Then oUserMarkColl.Add ElemsToRemark.Item(i)
    Next
    
    Dim oRemarkingSurface As IJSurfaceBody
    Set oRemarkingSurface = PartialPinJig.RemarkingSurface
       
    Dim oMfgMGHelper As New MfgMGHelper
    
    Dim oMarkLine As IJMfgMarkingLines_AE
    For Each oMarkLine In oUserMarkColl
        Dim oNI As IJNamedItem
        Set oNI = oMarkLine
        
        Dim oCSColl     As IJDObjectCollection
        Dim oProjCS     As IJComplexString
        Dim oCS         As IJComplexString
        
        Set oCSColl = oMarkLine.GeometryAsComplexStrings
        
        For Each oCS In oCSColl
    
            'Need to project the geometry onto the remarking surface, else might end up having lines extending beyond the contour
            On Error Resume Next
            oMfgMGHelper.ProjectComplexStringToSurface oCS, oRemarkingSurface, Nothing, oProjCS
            On Error GoTo ErrorHandler
            
            If oProjCS Is Nothing Then
                StrMfgLogError Err, MODULE, METHOD, "Failed to project marking line onto remarking surface", , , , "RULES"
                Set oProjCS = oCS
            End If
            
            CreateMfgGeom3dObject oProjCS, STRMFG_PinJigRemarkingLine3D, _
                                  oMarkLine, ReturnColl, oNI.Name, _
                                  oMarkLine.MfgMarkingType
    
            Set oCS = Nothing
            Set oProjCS = Nothing
            
        Next ' next oCS
        
    Next ' next oMarkLine
    
    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl
    
    Set ReturnColl = Nothing
    Set oMarkLine = Nothing
    Set oUserMarkColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5009, , "RULES")

End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
