Imports Ingr.SP3D.Common.Middle.Services
Public Class HSHelper
    Private Const UC_UserWarningMessage = "USERWARNINGMESSAGE"
    Private Const UC_UserErrorMessage = "USERERRORMESSAGE"

    Enum WarnOrError
        WarnOnly = 0
        ErrorOnly = 1
    End Enum

    Public Sub Notify(ByVal errorObject As ErrObject, ByVal description As String, ByVal method As String, ByVal sourceFile As String, Optional ByVal warnOrError As WarnOrError = WarnOrError.WarnOnly)
        Dim errorContext As String

        If (warnOrError = warnOrError.WarnOnly) Then
            errorContext = UC_UserWarningMessage
            MiddleServiceProvider.ErrorLogger.Log(errorObject.Number, "", "", errorContext, method & ": WARNING: " & description, "", sourceFile, 1)

        Else
            errorContext = UC_UserErrorMessage
            MiddleServiceProvider.ErrorLogger.Log(errorObject.Number, "", "", errorContext, method & ": ERROR: " & description, "", sourceFile, 1)
            Err.Raise(errorObject.Number)
        End If


    End Sub

End Class
