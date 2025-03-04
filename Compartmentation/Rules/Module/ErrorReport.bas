Attribute VB_Name = "ErrorReport"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    ErrorReport.bas
'
'  History:
'******************************************************************

Option Explicit


'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource    As String
    Dim strErrDesc      As String
    Dim lErrNumber      As Long
    Dim oEditErrors     As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file
    ' specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/
    '      ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

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
