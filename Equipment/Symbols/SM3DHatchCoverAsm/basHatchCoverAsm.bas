Attribute VB_Name = "basHatchCoverAsm"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basHatchCoverAsm.cls
'   Author: RUK
'   Creation Date:  Monday, July 28 2008
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   28.07.2008      RUK     Created
'******************************************************************************
Option Explicit

'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub



