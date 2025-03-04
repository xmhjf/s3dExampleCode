Attribute VB_Name = "basBHTy1Asm"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basBHTy1Asm.cls
'   Author:         dkl
'   Creation Date:  Wednesday, August 22 2007
'
'   Change History:
'   dd.mmm.yyyy     who            change description
'   -----------     ---           ------------------
'   22.Aug.2007     dkl         CR 123851, Created the symbol.
'******************************************************************************

Option Explicit

'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub



