VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideTopFlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Top Flange Side Process Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Implements IJDMfgProfileUpsideRule

Private Const MODULE = "MfgTemplateProcessProfile.SideTopFlg"

Private Function IJDMfgProfileUpsideRule_GetUpside(ByVal oStructProfilePart As Object) As Long
    Const METHOD = "UpsideTop: IJDMfgProfileUpsideRule_GetUpside"
    
    On Error GoTo ErrorHandler
   
    If TypeOf oStructProfilePart Is IJProfilePart Then
        IJDMfgProfileUpsideRule_GetUpside = CheckValidFaceTypeForProfile(oStructProfilePart, JXSEC_TOP)
    Else
        IJDMfgProfileUpsideRule_GetUpside = CheckValidFaceTypeForMember(oStructProfilePart, JXSEC_TOP)
    End If
   
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4018, , "RULES")
 
End Function

 
