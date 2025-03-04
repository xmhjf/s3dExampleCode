VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bottom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMarkingMember
' Module: Bottom
'
' Description:  Determines the marking side for the mfg member
'  History:
'       Suma Mallena         May 02. 2008   created
'
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileUpsideRule

Const MODULE = "MfgMarkingMember"

Private Function IJDMfgProfileUpsideRule_GetUpside(ByVal MemberPart As Object) As Long
    Const METHOD = "Bottom: IJDMfgProfileUpsideRule_GetUpside"
    
    On Error GoTo ErrorHandler
   
    IJDMfgProfileUpsideRule_GetUpside = CheckValidFaceTypeForMember(MemberPart, JXSEC_BOTTOM)
    
    Exit Function
    
ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4031, , "RULES")

End Function
 


