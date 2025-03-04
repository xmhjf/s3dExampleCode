VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamRemark_ButtPC"
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
Private Const MODULE As String = "MfgPinJigRule.SeamRemark_ButtPC::"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim oUniqueButtPCs As IJElements
    Set oUniqueButtPCs = GetConnectionsBetweenPinJigPlates(ConnectionPhysical, oPinJig)
    
    Dim oSeamColl As IJElements
    Dim Dummy1() As Long
    Dim Dummy2() As Long
    Dim Dummy3() As Long
    Dim Dummy4() As Long
    
    If oUniqueButtPCs.Count > 0 Then
        Set oSeamColl = GetOperatorsFromConnections(oUniqueButtPCs, Dummy1, Dummy2, Dummy3, Dummy4)
    Else
        Set oSeamColl = New JObjectCollection
    End If
    
    GetAPSMarkingLines oPinJig, STRMFG_SEAM_MARK, oSeamColl
    
    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = oSeamColl
    
    Set oPinJig = Nothing
    Set oUniqueButtPCs = Nothing
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
    
    Dim oUniqueButtPCs As IJElements
    Set oUniqueButtPCs = GetConnectionsBetweenPinJigPlates(ConnectionPhysical, oPinJig)
    
    Dim oSeamColl As IJElements
    Dim OprIdPerConn() As Long
    Dim OpnIdPerConn() As Long
    Dim UniqOprIds() As Long
    Dim UniqOpnIds() As Long
    
    If oUniqueButtPCs.Count > 0 Then
        Set oSeamColl = GetOperatorsFromConnections(oUniqueButtPCs, OprIdPerConn, OpnIdPerConn, UniqOprIds, UniqOpnIds)
    Else
        Set oSeamColl = New JObjectCollection
    End If
    
    Dim RemarkingSurface As IJSurfaceBody
    Set RemarkingSurface = oPinJig.RemarkingSurface
    
    Dim i As Long
    For i = 1 To oUniqueButtPCs.Count
        Dim oSeamObj As Object
        Set oSeamObj = oSeamColl.Item(LookUpUniqList(OprIdPerConn(i), OpnIdPerConn(i), UniqOprIds, UniqOpnIds))
        
        If ElemsToRemark.Contains(oSeamObj) Then
            Dim oNI As IJNamedItem
            Set oNI = oSeamObj
                
            Dim oCScoll As IJElements
            Set oCScoll = GetPCPortionOnSurface(oUniqueButtPCs.Item(i), RemarkingSurface)
            
            Dim oCS As IJComplexString
            For Each oCS In oCScoll
                CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, oSeamObj, _
                                      ReturnColl, oNI.Name, STRMFG_PinJig_Remarking_Seam
                Set oCS = Nothing
            Next
        End If
        
        Set oCScoll = Nothing
        Set oNI = Nothing
        Set oSeamObj = Nothing
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
    Set oUniqueButtPCs = Nothing
    Set RemarkingSurface = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
