VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckMfgPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: CCheckMfgPlate.cls
'
' Author: J.Manasa
'
' Abstract:
'   Check Manufacturability rule that checks the concurrency of Manufacturing plates
'   This rule checks if there is more than one Manufacturing plate for a plate part
'********************************************************************
Option Explicit

Private Const Module = "StrMfgCheckMfctyPlate.CCheckMfgPlate"
Private Const strProgID = "StrMfgCheckMfctyPlate.CCheckMfgPlate"

Private Const IID_IJStructGeometry = "{6034AD40-FA0B-11D1-B2FD-080036024603}"
Private Const IID_IJMfgPlateCreation_AE = "{C122691E-C84C-4d50-A04C-8BA7F6201CB9}"
Private Const DETAILPATEPARTINPUT_DEST = "DetailPlatePartInput_DEST"
Private Const MFGPLATEOUTPUT_ORIG = "MfgPlateOutput_ORIG"
Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"
Private Const IID_IJMfgPlatePart = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"

'localizer strings
Private m_ErrMsg As String
Private m_ViewLogFile As String
Private m_ProcessingCompleted As String
Private m_ErrDetected As String

Implements IJCheckMfcty

Public Enum ESeverityIndex
    siError = 101
    siWarning = 102
End Enum
Private Sub Class_Initialize()
    Const METHOD As String = "Class_Initialize"
    On Error GoTo ErrorHandler
    m_bLogError = False
    GoTo Cleanup
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
Cleanup:
End Sub

''////////////////////////////////////////////////////////////////////
'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description:
' This method is called by the Check Manufacturability mechanism
' to perform the desired checks for this rule.This method returns
' the plate parts which have more than one manufacturing plate
'
' Inputs:
'     oElements     :   The selected items in the working set or the entire working set.
'     bStopOnError  :   Operator option to stop after first error
'                       has been encountered.
'     pOptionCodes  :   not used in this method.
'     pCallingObj   :   The object that made the call to this method.
'                       This is used to pass back information to the Check
'                       Manufacturability command.
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck( _
        ByVal oElements As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        pOptionCodes() As Variant, _
        ByVal pCallingObj As Object)
Const METHOD = "IJCheckMfcty_Execute"
    On Error GoTo ErrorHandler

    Dim nCount As Long
    Dim nMaxCount As Long
        
    Dim lngDummy As Long
    Dim lngCancel As Long

    Dim strMessage As String
    Dim bErrorOccurred As Boolean

    Dim oPOM As IJDPOM
    Dim oObject As Object
    Dim vMoniker As Variant
    Dim oCallBack As IJCheckMfctyCallback
    
    InitializeLocalizeVariables
    
    ' Setup the Check Manufacturability call back mechanism.
    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    ' If the working set is empty, do not continue.
    If oElements.Count = 0 Then
        Exit Sub
    End If

    ' Get the Pesistent Object Manager.
    Set oPOM = GetPOMEx("Model")
      
    'loop
    nCount = 0
    nMaxCount = oElements.Count
    
    m_bLogError = False
    For Each vMoniker In oElements
        ' Get the object from the moniker.
        Set oObject = oPOM.GetObject(vMoniker)
        
        ' Check the values.
        bErrorOccurred = CheckObjectForErrorCondition(oObject, strMessage)
        
        If bErrorOccurred Then
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oObject                the object with error/warning
            '   strProgID              prog id of this rule
            '   ESeverity.siError      severity of the error
            '                          (101 error or 102 warning)
            '   106(1)                    code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            '   strAction              type of repair to be made
            '   strActionTool          progid of the COM object that will make the repair
            oCallBack.OnCheckError oObject, strProgID, _
                                   ESeverityIndex.siError, 1, strMessage, _
                                   "COM Repair Object", _
                                   "StrMfgCheckMfctyPlate.RepairMfgPlate"
            
            ' The operator may desire that only the first error be found.
            If bStopOnError Then
                Exit For
            End If
        End If
            
        ' Update the progress bar.  First argument is the progress as a
        ' percentage (i.e. 0 - 100).  The second argument is a cancellation
        ' flag.
        nCount = nCount + 1
        oCallBack.OnProgress nCount * 100 / nMaxCount, 0
    Next vMoniker

    If m_bLogError Then
        MsgBox m_ProcessingCompleted & vbCrLf & _
               m_ErrDetected & vbCrLf & _
               m_ViewLogFile, , _
               "StrMfgCheckMfctyPlate.CCheckMfgPlate"
    End If
    
    Set oCallBack = Nothing
    Set vMoniker = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
End Sub

'******************************************************************************
' Routine: CheckObjectForErrorCondition
'
' Abstract:
'  Checks if the current Object is a Plate part
'  If it is a plate part,get the number of manufacturing plates it has.
'  If there is a concurrency and it is not as a result of partial split then report it.
' Description:
'
' Inputs:
'
' Outputs:
'
'******************************************************************************
Private Function CheckObjectForErrorCondition(oObject As Object, _
                                              ByRef errorMessage As String) As Boolean
Const METHOD As String = "CheckObjectForErrorCondition"
    On Error GoTo ErrorHandler

    Dim sMessage As String
    Dim oAllMfgPlateParts As IJElements
    Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    
    CheckObjectForErrorCondition = False
        
    ' Check if current object is type IJPlatePart
    If TypeOf oObject Is IJPlatePart Then
       
        'Get all mfgPlatePart under this Plate part
        Set oAllMfgPlateParts = oStructMfgGlobals.GetMfgPart(oObject, IID_IJMfgPlatePart)
        
        If Not oAllMfgPlateParts Is Nothing Then
            If oAllMfgPlateParts.Count > 1 Then
                    CheckObjectForErrorCondition = True
                    errorMessage = m_ErrMsg
            End If
        End If
    End If
    
    Set oAllMfgPlateParts = Nothing
    Set oStructMfgGlobals = Nothing
   Exit Function
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Function

Private Sub InitializeLocalizeVariables()
    Const METHOD = "InitializeLocalizeVariables"
    On Error GoTo ErrorHandler
    
    Dim sMessage As String
    
    Dim oLocalizer As IJLocalizer
    Dim oContext As IJContext
    Set oContext = GetJContext()
    Const CONTEXTSTRING = "OLE_SERVER"
    
    Set oLocalizer = InitializeLocalizer(oContext.GetVariable(CONTEXTSTRING) & "\bin\StructManufacturing\Resource\", _
                                        App.EXEName)
    
    m_ErrMsg = oLocalizer.GetString(IDS_CHECKMFGPLATE_MSG_ERRORMESSAGE, "The Platepart has invalid number of Manufacturing plates")
    m_ViewLogFile = oLocalizer.GetString(IDS_CHECKMFGPLATE_MSG_VIEWLOGFILE, "see Sp3d ErrorLog for Manufacturing")
    m_ProcessingCompleted = oLocalizer.GetString(IDS_CHECKMFGPLATE_MSG_PROCESSINGCOMPLETE, "Processing Complete")
    m_ErrDetected = oLocalizer.GetString(IDS_CHECKMFGPLATE_MSG_ERRDETECTED, "Errors (and/or Warnings) were detected")
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Sub
