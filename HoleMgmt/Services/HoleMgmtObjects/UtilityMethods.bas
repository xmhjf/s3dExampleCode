Attribute VB_Name = "UtilityMethods"
'******************************************************************************
' Copyright (C) 1998-2006 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\HoleMgmtObjects
'
' File: UtilityMethods.bas
'
' Author: Hole Mgmt Team
'
' Abstract: utility methods used by HoleMgmtObjects
'******************************************************************************

Option Explicit

Private Const sSOURCEFILE As String = "HoleMgmtObjects::UtilityMethods.bas"

'******************************************************************************
' Routine: LogError
'
' Abstract: log an error to the error service
'******************************************************************************
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
       
    ' add the error to the service : the error is also logged to the file specified by
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

'******************************************************************************
' Routine: GetResourceMgrFromObject
'
' Abstract: return the resource manager from an object
'******************************************************************************
Public Function GetResourceMgrFromObject(ByVal pObject As Object) As Object
    On Error Resume Next
    
    Dim oJDObject As IJDObject
    Dim oResourceMgr As IUnknown
    
    Set oJDObject = pObject
    If Not oJDObject Is Nothing Then
        Set oResourceMgr = oJDObject.ResourceManager
    
        If Not oResourceMgr Is Nothing Then
            Set GetResourceMgrFromObject = oResourceMgr
            Set oResourceMgr = Nothing
        End If
        
        Set oJDObject = Nothing
    End If
        
    Exit Function
End Function
 