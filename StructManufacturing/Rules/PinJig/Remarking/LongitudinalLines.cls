VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongitudinalLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: FrameLines
'
' Description:  Example to show the correct longitudinal lines on the pinjig
'
' Author: H.S. Jeong
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.LongitudinalLines"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IMSCoreCollections.IJElements
    Const METHOD As String = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim oDirectionVec As DVector
    Set oDirectionVec = New DVector

    ' Longitudinal lines => Plane's normal is along Y direction
    oDirectionVec.Set 0, 1, 0

    ' This is the collection that will be returned
    Dim oFrameSet As IMSCoreCollections.IJElements
    Set oFrameSet = GetReferencePlanesInPinJigRange(PartialPinJig, oDirectionVec, 5014)

    ' Specific hack for L0
    Dim xLow As Double, yLow As Double, zLow As Double
    Dim xHigh As Double, yHigh As Double, zHigh As Double
    
    Dim oRange As IJRangeAlias
    Set oRange = PartialPinJig
    
    Dim oGBox As GBox
    oGBox = oRange.GetRange
    
    xLow = oGBox.m_low.X
    yLow = oGBox.m_low.Y
    zLow = oGBox.m_low.Z
    
    xHigh = oGBox.m_high.X
    yHigh = oGBox.m_high.Y
    zHigh = oGBox.m_high.Z
    
    If yLow < 0# And yHigh > 0# Then
        Dim BOwithFS As IJDMfgFrameSystem
        Set BOwithFS = PartialPinJig
        
        Dim oPinJigFrameSys As IHFrameSystem
        Set oPinJigFrameSys = BOwithFS.FrameSysParent
        
        If oPinJigFrameSys Is Nothing Then GoTo CleanUp
        
        Dim oCenterLineFrame As IHFrame
        Set oCenterLineFrame = oPinJigFrameSys.PrincipalYAxis.Frame("L0")
        
        If Not oCenterLineFrame Is Nothing Then
            oFrameSet.Add oCenterLineFrame
        End If
        
        Set oCenterLineFrame = Nothing
    End If
    ' Send the data back
    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = oFrameSet
CleanUp:
    Set oDirectionVec = Nothing
    Set oFrameSet = Nothing
    Set oPinJigFrameSys = Nothing
    Set BOwithFS = Nothing
    Set oPinJigFrameSys = Nothing
    Set oCenterLineFrame = Nothing
    Set oRange = Nothing
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5014, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IMSCoreCollections.IJElements) As String
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsAnYPlane]"
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IMSCoreCollections.IJElements) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
Const METHOD As String = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    Dim ReturnColl As IJMfgGeomCol3d
    
    CreateFrameRemarkingLines PartialPinJig, ElemsToRemark, STRMFG_PinJig_Remark_FrameY, ReturnColl
    
        Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl
        
        Set ReturnColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5015, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
