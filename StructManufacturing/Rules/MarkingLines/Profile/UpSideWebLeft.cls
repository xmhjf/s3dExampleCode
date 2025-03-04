VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideWebLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMarkingProfile
' Module: UpsideWebLeft
'
' Description:   Determines the marking side for the mfg profile
'  History:
'       Suma Mallena         May 02. 2008   created
'
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileUpsideRule

Const MODULE = "MfgMarkingProfile"

Private Function IJDMfgProfileUpsideRule_GetUpside(ByVal oStructProfilePart As Object) As Long
    Const METHOD = "UpsideWebLeft: IJDMfgProfileUpsideRule_GetUpside"
    
    On Error GoTo ErrorHandler
    
    If TypeOf oStructProfilePart Is IJProfilePart Then
        IJDMfgProfileUpsideRule_GetUpside = CheckValidFaceTypeForProfile(oStructProfilePart, JXSEC_WEB_LEFT)
    Else
        IJDMfgProfileUpsideRule_GetUpside = CheckValidFaceTypeForMember(oStructProfilePart, JXSEC_WEB_LEFT)
    End If
    
    Exit Function
    
ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4016, , "RULES")

End Function
 
