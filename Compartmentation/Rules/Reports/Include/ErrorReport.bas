Attribute VB_Name = "ErrorReport"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    ErrorReport.bas
'
'  Initial Creation : Thandur Raghuveer
'
'  History:
'******************************************************************

Option Explicit


'********************************************************************
' ' Routine: CompartLogError
'
' Description:  Compartmentation specific Error logger
'********************************************************************
Public Function CompartLogError(oErrObject As ErrObject, _
                               strModule As String, _
                               strMethod As String, _
                               Optional strExtraInfo As String = "", _
                               Optional strCodelistName As String = "CompartCustomErrorMessages", _
                               Optional lCodelistNum As Long, _
                               Optional strItemName As String = "", _
                               Optional strErrCtxSfx As String = "") As Long
                               
    Dim oCompartErrorLog As IJDCompartErrorLog
    Set oCompartErrorLog = New CompartErrorLog
    
    If oCompartErrorLog Is Nothing Then Exit Function
    
    oCompartErrorLog.LogError oErrObject, _
                           strItemName, _
                           strModule, _
                           strMethod, _
                           strCodelistName, _
                           lCodelistNum, _
                           strExtraInfo, _
                           strErrCtxSfx
    CompartLogError = oErrObject.Number
End Function
