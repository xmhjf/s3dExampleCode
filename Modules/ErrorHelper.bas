Attribute VB_Name = "ErrorHelper"
Option Explicit

Private Const MODULE = "ErrorHelper:"  'Used for error messages

Public Const E_FAIL = -2147467259

Function ReportError(pErrObject As ErrObject, Optional strSourceFile As String = "", Optional strMethod As String = "", Optional strDataSource As String = "", Optional lLineNumber As Long = 0) As IJError
   
     ' retrieve the error service
    Dim pEditErrors As IJEditErrors
    Set pEditErrors = GetJContext.GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set ReportError = pEditErrors.Add(pErrObject.Number + vbObjectError, pErrObject.Source, pErrObject.Description, "", "", 0, _
                      "", strDataSource, strSourceFile & "::" & strMethod, lLineNumber)
End Function
'
' log the error into the error collection
'
' the last error is kept in the error holder, so that the calling semantic can supply extra info
' the error folder is pushed into the error collection, when a new error is logged
' the extrainfo will be supplied by the calling semantic, when returning from VB code
'
Sub LogError(ByRef pErrObject As ErrObject, Optional sDataSource As String = "", Optional sSourceFile As String = "", Optional lSourceLine As Long = 0)
    'declare that the error is returned by VB
    Dim lBasicErrorNumber As Long
    Let lBasicErrorNumber = &H80000000
    
    If (pErrObject.Number And lBasicErrorNumber) = 0 Then
        Let pErrObject.Number = pErrObject.Number Or vbObjectError 'add &H80040000
    End If
   
    Dim pEditErrors As IJEditErrors
    ' retrieve the error service
    ' errors persist in the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set pEditErrors = GetJContext.GetService("Errors")
    
    Call pEditErrors.AddFromErr(pErrObject, "", sDataSource, sSourceFile, lSourceLine)
End Sub
'
' log error into the error collection and raise the error to propagate it back to the calling routine
'
Sub LogAndRaiseError(ByRef pErrObject As ErrObject, Optional sDataSource As String = "", Optional sSourceFile As String = "", Optional lSourceLine As Long = 0)
    ' log the error
    Call LogError(pErrObject, sDataSource, sSourceFile, lSourceLine)
    
    ' re-raise the error to propagate it back to the calling routine
    Call Err.Raise(pErrObject.Number, pErrObject.Source, pErrObject.Description, pErrObject.HelpFile, pErrObject.HelpContext)
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
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "ErrorHelper.bas": Let sMethod = "SPSToDoErrorNotify"
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
    LogError Err, sSourceFile, sMethod
    Err.Clear
End Sub

'*************************************************************************
'Function
'SPSToDoErrorNotifyEx
'
'Abstract
' Called to notify the SmartOccurrence of a ToDo localized error that occurred during a
' smart occurrence custom evaluate
'
'History
'
'***************************************************************************
Public Sub SPSToDoErrorNotifyEx(strModuleName As String, nMessageNumber As Long, strMessageText As String, oObjectInError As Object, oObjectToUpdate As Object)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "ErrorHelper.bas": Let sMethod = "SPSToDoErrorNotify"

    Dim oLocalizer As IJLocalizer: Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & strModuleName

    Dim strLocalizedError As String: strLocalizedError = oLocalizer.GetString(nMessageNumber, strMessageText)

    On Error GoTo ErrHandler:
    Dim oToDoListHelper As IJToDoListHelper ' Set ToDoListHelper = pointer to the CAO Object

    Set oToDoListHelper = oObjectInError
    If Not oToDoListHelper Is Nothing Then
        If oObjectToUpdate Is Nothing Then
          oToDoListHelper.SetErrorInfoEx strModuleName, nMessageNumber, strMessageText, ErrorTypeError
        Else
          oToDoListHelper.SetErrorInfoEx strModuleName, nMessageNumber, strMessageText, ErrorTypeError, oObjectToUpdate
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    ' Report the error but do not return the error since this call maybe in their
    '   error handler
    LogError Err, sSourceFile, sMethod
    Err.Clear
End Sub



