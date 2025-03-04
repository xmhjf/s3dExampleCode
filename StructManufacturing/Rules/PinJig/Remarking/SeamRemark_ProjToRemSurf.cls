VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamRemark_ProjToRemSurf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
'
' Author: Anand Hariharan
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.SeamRemark_ProjToRemSurf::"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim oUniqueLogConns As IJElements
    Set oUniqueLogConns = GetConnectionsBetweenPinJigPlates(ConnectionLogical, oPinJig)
    
    Dim oSeamColl As IJElements
    Dim Dummy1() As Long
    Dim Dummy2() As Long
    Dim Dummy3() As Long
    Dim Dummy4() As Long
    
    If oUniqueLogConns.Count > 0 Then
        Set oSeamColl = GetOperatorsFromConnections(oUniqueLogConns, Dummy1, Dummy2, Dummy3, Dummy4)
    Else
        Set oSeamColl = New JObjectCollection
    End If
    
    GetAPSMarkingLines oPinJig, STRMFG_SEAM_MARK, oSeamColl
    
    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = oSeamColl
    
    Set oPinJig = Nothing
    Set oUniqueLogConns = Nothing
    Set oSeamColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities"
    On Error GoTo ErrorHandler
    
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsDesignSeam]"
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler
    
    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim oSeamColl As IJElements
    Set oSeamColl = New JObjectCollection
    
    Dim i As Long
    For i = 1 To ElemsToRemark.Count
        Dim IthElemToRemark As Object
        Set IthElemToRemark = ElemsToRemark.Item(i)
        If TypeOf IthElemToRemark Is IJWireBody And _
           TypeOf IthElemToRemark Is IJNamedItem Then
             oSeamColl.Add ElemsToRemark.Item(i)
        End If
    Next

    Dim RemarkingSurface As IJSurfaceBody
    Set RemarkingSurface = oPinJig.RemarkingSurface
    
    Dim oMfgMGhelper As New MfgMGHelper
    
    ' Now project the seams on the remarking surface.
    For i = 1 To oSeamColl.Count
        Dim oSeamGeom As IJWireBody
        Set oSeamGeom = oSeamColl.Item(i)
        
        Dim oNI As IJNamedItem
        Set oNI = oSeamColl.Item(i)
        
        Dim oSeamCS As IJComplexString
        oMfgMGhelper.WireBodyToComplexString oSeamGeom, oSeamCS
        
        Dim oCScoll As IJElements
        oMfgMGhelper.ProjectCSToSurface oSeamCS, RemarkingSurface, Nothing, oCScoll
        
        Dim oCS As IJComplexString
        For Each oCS In oCScoll
            CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, oSeamColl.Item(i), _
                                  ReturnColl, oNI.Name, STRMFG_PinJig_Remarking_Seam
            Set oCS = Nothing
        Next
        
        Set oSeamGeom = Nothing
        Set oNI = Nothing
        Set oSeamCS = Nothing
        Set oCScoll = Nothing
    Next
    
    Dim oSeamMarks As IJElements
    Set oSeamMarks = New JObjectCollection
     
    GetAPSMarkingLines oPinJig, STRMFG_SEAM_MARK, oSeamMarks

    If oSeamMarks.Count > 0 Then
        CreateGeom3dFromAPSMarkingLines oPinJig, ElemsToRemark, oSeamMarks, STRMFG_PinJig_Remarking_Seam, ReturnColl
    End If
    Set oSeamMarks = Nothing
    
    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl
    
CleanUp:
    Set ReturnColl = Nothing
    Set oPinJig = Nothing
    Set oSeamColl = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function


