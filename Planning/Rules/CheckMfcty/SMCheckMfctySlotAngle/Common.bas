Attribute VB_Name = "Common"

Public Const CUSTOMERID = "SM"
Public Const m_sProjectName As String = CUSTOMERID + "CheckMfctySlotAngle"
Public Const m_sProjectPath As String = "S:\StructDetail\Data\" + m_sProjectName
Public Const IID_IJFullObject As String = "{bcbfb3c0-98c2-11d1-93de-08003670a902}"

Private Const m_sModule As String = "Common.bas"

Public Function LogError(ByVal oErrObject As ErrObject, _
                         Optional ByVal strSourceFile As String = "", _
                         Optional ByVal strMethod As String = "", _
                         Optional ByVal strExtraInfo As String = "") As IJError
    If oErrObject Is Nothing Then
       Set LogError = Nothing
       Exit Function
    End If
    
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
    
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' Add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
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

' Method: GetPOM
'
' This method returns the Persistent Object Manager (POM) for the specified DB
'
Public Function GetPOM(sDatabaseType As String) As IJDPOM
    Const sMethod As String = "GetPOM"
    On Error GoTo ErrorHandler
        
    ' Get the Server Context.
    Dim oPOM As IJDPOM
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle

    Set oContext = GetJContext()

    ' Get the AccessMiddle object.
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oContext = Nothing

    ' Get the Persistent Object Managers (POM).
    Set oPOM = oAccessMiddle.GetResourceManagerFromType(sDatabaseType)
    Set oAccessMiddle = Nothing

    ' Return the POM.
    Set GetPOM = oPOM
   
    Exit Function

ErrorHandler:
   Err.Raise LogError(Err, m_sModule, sMethod, "Can not get POM").Number

End Function

