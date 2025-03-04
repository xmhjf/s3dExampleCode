Attribute VB_Name = "OH2HorizontalPump"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   OH2HorizontalPump.bas
'   Author: Veena
'   Creation Date:  Monday, Jan 08 2006
'
'   Description:
'       TODO - fill in header description information
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

