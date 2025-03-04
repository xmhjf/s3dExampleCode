Attribute VB_Name = "SECommon"
Option Explicit
Private Const MODULE = "CustomReports.SECommon : "

Public Function GetAngle(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetAngle"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Set oTemplateRpt = oTemplate
    GetAngle = oTemplateRpt.GetTemplateAttachedAngle
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Public Function GetFrames(oTemplateSet) As IJElements
    Const METHOD As String = "GetFrames"
    On Error GoTo ErrorHandler
    
    Dim oSettingsHelper  As MfgSettingsHelper
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim oTemplateSetRpt As IJMfgTemplateSetReport

    Dim strProgId As String
    Dim oPlatePart As IJPlatePart
    Dim oPositionFrames As IJMfgTemplatePositionFrameRule
    
    Set oProcessSettings = oTemplateSet.GetProcessSettings
    Set oSettingsHelper = oProcessSettings
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("PositionFrames")

    Set oPositionFrames = SP3DCreateObject(strProgId)
    
    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart
    
    Dim oFrameSystem As IHFrameSystem
    Dim oMfgFrameSys As IJDMfgFrameSystem
    
    Set oMfgFrameSys = oTemplateSet
    Set oFrameSystem = oMfgFrameSys.FrameSysParent
    Set GetFrames = oPositionFrames.GetPositionFrame(oFrameSystem, oPlatePart, oProcessSettings, oTemplateSet)
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Public Function IsBaseSide(oTemplateSet As IJDMfgTemplateSet) As Boolean
    Const METHOD As String = "IsBaseSide"
    On Error GoTo ErrorHandler
    
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim oPlateSideRule   As IJDMfgPlateUpSideRule
    Dim Upside As enumPlateSide
    Dim oSettingsHelper As MfgSettingsHelper
    Dim strProgId As String
    
    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Dim oPlatePart As IJPlatePart
    
    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart
    
    Set oProcessSettings = oTemplateSet.GetProcessSettings
    
    Set oSettingsHelper = oProcessSettings
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    
    Set oPlateSideRule = SP3DCreateObject(strProgId)
    Upside = oPlateSideRule.GetPlateUpSide(oPlatePart)

    If (Upside = BaseSide) Then
        IsBaseSide = True
    ElseIf (Upside = OffsetSide) Then
        IsBaseSide = False
    End If

    Set oProcessSettings = Nothing
    Set oSettingsHelper = Nothing
    Set oPlateSideRule = Nothing
    Set oTemplateSetRpt = Nothing
    Set oPlatePart = Nothing
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Public Function GetAftButtName(oTemplateSet As IJDMfgTemplateSet) As String
    Const METHOD As String = "GetAftButtName"
    On Error GoTo ErrorHandler
    
    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Dim oPlatePart As IJPlatePart
    
    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart
    
    GetAftButtName = oTemplateSetRpt.GetAftButtName(oPlatePart, IsBaseSide(oTemplateSet))

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Public Function GetForeButtName(oTemplateSet As IJDMfgTemplateSet) As String
    Const METHOD As String = "GetForeButtName"
    On Error GoTo ErrorHandler
    
    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Dim oPlatePart As IJPlatePart
    
    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart
    
    GetForeButtName = oTemplateSetRpt.GetForeButtName(oPlatePart, IsBaseSide(oTemplateSet))

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function
 
