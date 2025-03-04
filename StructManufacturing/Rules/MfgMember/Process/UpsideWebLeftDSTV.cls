VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideWebLeftDSTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2013, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: UpsideWebLeftDSTV
'
' Description:  Determines the WebLeft as upside for member and sets the MemberWeb to full height.
'
' Author:
'
' Comments:
'
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileUpsideRule

Const MODULE = "MfgProfileProcess"

Private Function IJDMfgProfileUpsideRule_GetUpside(ByVal ProfilePart As Object) As Long
    Const METHOD = "UpsideWebLeft: IJDMfgProfileUpsideRule_GetUpside"
    On Error GoTo ErrorHandler
    
        Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
        Set oProfileWrapper.object = ProfilePart
    
        Dim oMfgEntityHelper As IJMfgEntityHelper
        Set oMfgEntityHelper = New MfgEntityHelper
        Dim oMfgPart As IJMfgProfilePart
        If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
            oMfgEntityHelper.ReadXMLForPartUpside oMfgPart, "webleft", GetSymbolSharePath & UPSIDE_XML_LOCATION
            oMfgPart.ExtendMemberWebToFullHeight = True
        End If
        
        Set oProfileWrapper = Nothing
        IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_LEFT
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2032, , "RULES")
End Function

