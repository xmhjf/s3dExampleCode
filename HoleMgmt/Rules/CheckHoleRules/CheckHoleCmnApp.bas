Attribute VB_Name = "CheckHoleCmnApp"
'*******************************************************************
'  Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADPinPtCmd
'  File:    \CommonApp\Client\Modules\GSCADCmnAppSub.bas
'
'  Abstract: The file contains common subroutines and functions for GSCAD apps
'
'  Description: This modules contains common subroutines and functions that can be
'               used by GSCAD Commands.  The following is a list of the functions
'               provided by this module
'
'       InitGuid                    - Initializes a GUID structure
'       ReportUnanticipatedError    - Calls the IMSErrorService object for the current err
'
'
'  Author: Jim Fleming
'
'  History:
'     jpf   06-22-98    Initial release
'
'******************************************************************

Option Explicit

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
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

''//////////////////////////////////////////////////////////////////////////
'' InitGuid
''      Initialize a GUID
''//////////////////////////////////////////////////////////////////////////
Public Function InitGuid(a As Byte, b As Byte, c As Byte, d As Byte, e As Byte, f As Byte, _
                          g As Byte, h As Byte, i As Byte, j As Byte, k As Byte, l As Byte, _
                          m As Byte, n As Byte, o As Byte, p As Byte) As Variant

    Dim Guid(0 To 15) As Byte
    
    '' NOTE that the bytes are not in straight order : they're reversed
    '' based on the GUID 4byte/2byte/2byte start
    Guid(0) = d:  Guid(1) = c:  Guid(2) = b:  Guid(3) = a:
    Guid(4) = f:  Guid(5) = e:  Guid(6) = h:  Guid(7) = g:
    Guid(8) = i:  Guid(9) = j:  Guid(10) = k: Guid(11) = l:
    Guid(12) = m: Guid(13) = n: Guid(14) = o: Guid(15) = p
    
    InitGuid = Guid
End Function
 