Attribute VB_Name = "HgrSupRulesModule"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   HgrSupRulesModule.bas
'   ProgID:
'   Author:         Srinivas
'   Creation Date:  03.Mar.2005
'   Description:
'
'
'   Change History:
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
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

' ---------------------------------------------------------------------------
' Name: LogErrorWithContext
' Description: default Error logger, same as LogError but adds the errors with proper error context
'
' Inputs - oErrObject - ErrObject, [strSourceFile] - string, [strMethod] - string, [strExtraInfo] - string
' Outputs - IJError
' ---------------------------------------------------------------------------
Public Function LogErrorWithContext(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError

Const METHOD = "LogErrorWithContext"
On Error GoTo ErrHandler

    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors

    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description

     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")

    ' add the error to the service : the error is also logged to the file
    ' specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/
    '      ReportErrors_Log"
    Set LogErrorWithContext = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      UC_UserErrorMessage, _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
Exit Function
ErrHandler:
    Err.Raise Err.Number
End Function
