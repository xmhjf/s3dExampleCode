Attribute VB_Name = "basWearPlate"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basWearPlate.bas
'   Author: Huo
'   Creation Date:  November 2007
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

