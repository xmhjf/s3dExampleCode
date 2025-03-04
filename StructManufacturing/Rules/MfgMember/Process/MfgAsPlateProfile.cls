VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgAsPlateProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module:MfgAsProfileYes
'
' Description:  Provides the MfgAsPlate value for the mfg profile
'
' Author:
'
' Comments:
' 29 Aug, 2011    kkamph    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileProcessMfgAsPlateRule
Const MODULE = "MfgProfileProcess"

Private Function IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate(ByVal ProfilePart As Object) As Long
    Const METHOD = "MfgAsPlate: IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate"
    On Error GoTo ErrorHandler
        Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
        Set oProfileWrapper.object = ProfilePart
    
        Dim oMfgEntityHelper As IJMfgEntityHelper
        Set oMfgEntityHelper = New MfgEntityHelper
        Dim oMfgPart As IJMfgProfilePart
        If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
            oMfgEntityHelper.ReadXMLForPartOutput oMfgPart, "Plate-Profile", GetSymbolSharePath & UPSIDE_XML_LOCATION
        End If
    
        Set oProfileWrapper = Nothing
        SetBuiltUpNames ProfilePart, True
        
        IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate = 2
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2046, , "RULES")
End Function
 


