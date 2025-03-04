Attribute VB_Name = "ErrorReport"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    ErrorReport.bas
'
'  History:
'       MJV         april 14, 2004  Modified the error handling
'******************************************************************

Option Explicit

'********************************************************************
' ' Routine: LogError
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                         Optional strSourceFile As String = "", _
                         Optional strMethod As String = "", _
                         Optional strExtraInfo As String = "", _
                         Optional strCodelistName As String = "StrMfgSemanticErrorMessages", _
                         Optional lCodelistNum As Long = 7, _
                         Optional strItemName As String = "", _
                         Optional strErrCtxSfx As String = "") As IJError
    
    ' retrieve the error service
    Dim oEditErrors As IJEditErrors
    Set oEditErrors = GetJContext().GetService("Errors")
    
    Dim oEditError As IJEditError
    Set oEditError = oEditErrors.AddFromErr(oErrObject, _
                                            strExtraInfo, _
                                            strItemName, _
                                            strSourceFile)
    
    oEditError.Description = strMethod
    oEditError.HelpFile = strCodelistName
    oEditError.HelpContext = lCodelistNum
    oEditError.ErrorContext = strErrCtxSfx
    
    Set LogError = oEditError
    Set oEditErrors = Nothing
    Set oEditError = Nothing
End Function

'********************************************************************
' ' Routine: StrMfgLogError
'
' Description:  Struct Manufacturing specific Error logger
'********************************************************************
Public Function StrMfgLogError(oErrObject As ErrObject, _
                               strModule As String, _
                               strMethod As String, _
                               Optional strExtraInfo As String = "", _
                               Optional strCodelistName As String = "StrMfgSemanticErrorMessages", _
                               Optional lCodelistNum As Long = 7, _
                               Optional strItemName As String = "", _
                               Optional strErrCtxSfx As String = "") As Long
    
    Dim oStrMfgErrLog As IJDStrMfgErrorLog
    Set oStrMfgErrLog = New StrMfgErrorLog
    
    oStrMfgErrLog.LogError oErrObject, _
                           strItemName, _
                           strModule, _
                           strMethod, _
                           strCodelistName, _
                           lCodelistNum, _
                           strExtraInfo, _
                           strErrCtxSfx
    
    StrMfgLogError = oErrObject.Number
End Function
'********************************************************************
' Routine: LogMessage
'
' Description:  generic Error logger with out context.
'               If we log error with this method then SMEL_CHK_WARN_AND_END
'               doesn't put the object in the To Do List
'********************************************************************
Public Function LogMessage(oErrObject As ErrObject, _
                               strModule As String, _
                               strMethod As String, _
                               Optional strExtraInfo As String = "", _
                               Optional strCodelistName As String = "StrMfgSemanticErrorMessages", _
                               Optional lCodelistNum As Long = 7, _
                               Optional strItemName As String = "", _
                               Optional strErrCtxSfx As String = "") As Long
    
    Dim oContext As IJContext
    Set oContext = GetJContext()
    Dim oEditErrors As IJEditErrors
    Set oEditErrors = oContext.GetService("Errors")
    
    Dim oEditError As IJEditError
    Set oEditError = oEditErrors.AddFromErr(Err, strExtraInfo, strMethod, strModule)
           
    'Overwrite error with input codelist, method
    
    oEditError.HelpFile = strCodelistName
    oEditError.HelpContext = lCodelistNum
    oEditError.Description = strMethod
    
    LogMessage = oErrObject.Number
End Function
