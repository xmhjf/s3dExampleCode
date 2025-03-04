Attribute VB_Name = "CommonError"
Option Explicit

'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : StructCommonSym.bas
'
'Author : Aniket Patil
'
'Description :
'    SmartPlant Structural common Error Functions
'
''********************************************************************


'*************************************************************************
'Function
'ReportError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and returns an error code. The custom assembly code then put the objectt that
'raised the error to ToDo list
'
'input
'error object, file name where error occurred, method name, data source, line number
'
'Return
'error code
'
'Exceptions
'
'***************************************************************************
Public Const SPS_MACRO_WARNING = &HC0000000

Public Function ReportError(pErrObject As ErrObject, Optional strSourceFile As String = "", Optional strMethod As String = "", Optional strDataSource As String = "", Optional lLineNumber As Long = 0) As IJError
 
     ' retrieve the error service
    Dim pEditErrors As IJEditErrors
    Set pEditErrors = GetJContext.GetService("Errors")
    pErrObject.Description = strDataSource
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
  
    Set ReportError = pEditErrors.Add(pErrObject.Number, pErrObject.Source, pErrObject.Description, "", "", 0, _
                      "", strDataSource, strSourceFile & "::" & strMethod, lLineNumber)
    
End Function

'*************************************************************************
'Function
'HandleError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and returns success
'
'input
'Module(file) name, method name
'
'Return
'success
'
'Exceptions
'
'***************************************************************************
Public Sub HandleError(sModule As String, sMethod As String)
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    Set oEditErrors = Nothing
End Sub


'*************************************************************************
'Function
'SPSToDoErrorNotify
'
'Abstract
' Called to notify the SmartOccurrence of a ToDo error that occurred during a
' smart occurrence custom evaluate
'
'History
'
'***************************************************************************
Public Sub SPSToDoErrorNotify(strCodelistTable As String, nToDoListErrorNum As Long, oObjectInError As Object, oObjectToUpdate As Object)
    Const METHOD = "SPSToDoErrorNotify"
    On Error GoTo ErrHandler:
    Dim oToDoListHelper As IJToDoListHelper ' Set ToDoListHelper = pointer to the CAO Object

    Set oToDoListHelper = oObjectInError
    If Not oToDoListHelper Is Nothing Then
        If oObjectToUpdate Is Nothing Then
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum
        Else
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum, oObjectToUpdate
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    ' Report the error but do not return the error since this call maybe in their
    '   error handler
    HandleError "CommonError", METHOD
    Err.Clear
End Sub

Public Sub CheckForUndefinedValueAndRaiseError(oObject As Object, lAttributeValue As Long, strAttributeCodeListTableName As String, lErrorNumber As Long, Optional oObjectToUpdate As Object)
' No error handler, so that caller will get it.
    Dim oToDoListHelper As IJToDoListHelper
    If lAttributeValue = -1 Or (DoesCodelistValueExist(lAttributeValue, strAttributeCodeListTableName) = False) Then
        If Not oObject Is Nothing Then
            If TypeOf oObject Is IJDRepresentationDuringGame Then
                Set oToDoListHelper = oObject.Definition.IJDDefinitionPlayerEx.PlayingSymbol
            ElseIf TypeOf oObject Is IJToDoListHelper Then
                Set oToDoListHelper = oObject
            End If
            If Not oToDoListHelper Is Nothing Then
                oToDoListHelper.SetErrorInfo "SPSTDLCodeListErrors", lErrorNumber, oObjectToUpdate
                Err.Description = "Undefined Value"
                Err.Raise SPS_MACRO_WARNING
            End If
        End If
    End If
End Sub
