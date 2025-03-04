VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OffsetFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2012, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgTemplateProcessEdge
' Module:  OffsetFull
'
' Description:  Determines the process settings for the Edge Template offset, which will be applied to find location of bottom curve
'
' History:
'   20th April 2012    Siva     Creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateProcessEdge.OffsetFull"

Implements IJDMfgEdgeTemplateOffsetRule

' Defines the Offset value to be used for computing the template bottom line location
Private Function IJDMfgEdgeTemplateOffsetRule_GetTemplateOffset(ByVal pDispPart As Object, Optional ByVal lSectionID As Long) As Double
Const METHOD = "IJMfgEdgeTemplateExtensionRule_GetTemplateExtension"
On Error GoTo ErrorHandler
    
    ' Check for valid section ID and also check input part is profile part
    If lSectionID = 0 And Not TypeOf pDispPart Is IJProfilePart Then
        Exit Function
    End If
    
    Dim oProfilePartHelper  As MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHelper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfilePartHelper.object = pDispPart
    
    Dim oSurfacePort    As IJPort
    Set oSurfacePort = oProfilePartHelper.GetSurfacePort(lSectionID)
    
    Const strStandAloneProfilePartTrimAEProgId = "ProfilePartActiveEntities.ProfileTrim_AE.1"
    
    Dim oStructPort As IJStructPortEx
    Set oStructPort = oSurfacePort
    
    Dim oPortBeforeTrim As IJPort
    Set oPortBeforeTrim = oStructPort.RelatedPort(strStandAloneProfilePartTrimAEProgId, True)
    
    Dim oMfgEntityHelper As IJMfgEntityHelper
    Set oMfgEntityHelper = New MfgEntityHelper
    
    Dim oBaseEdgePortGeom   As Object
    Set oBaseEdgePortGeom = oMfgEntityHelper.GetBaseEdgeOfFace(oPortBeforeTrim)
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    Dim oBaseEdgeCurve As IJCurve
    Set oBaseEdgeCurve = oMfgRuleHelper.WireBodyToComplexString(oBaseEdgePortGeom)
    
    IJDMfgEdgeTemplateOffsetRule_GetTemplateOffset = oBaseEdgeCurve.length
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9999, , "RULES")
End Function

