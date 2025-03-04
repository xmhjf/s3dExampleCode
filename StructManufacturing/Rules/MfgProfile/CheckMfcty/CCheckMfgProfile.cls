VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: CCheckMfgProfile.cls
'
' Author: J.Manasa
'
' Abstract:
'   Check Manufacturability rule that checks the concurrency of Manufacturing profiles
'   This rule checks if there is more than one Manufacturing profile for a profile part
'********************************************************************
Option Explicit

Private Const Module = "StrMfgCheckMfctyProfile.CCheckProfile"
Private Const strProgID = "StrMfgCheckMfctyProfile.CCheckProfile"

Private Const IID_IJStructProfilePart = "{F933A05F-8820-4738-B743-E030EC9502AB}"
Private Const IID_IJMfgProfileCreation_AE = "{6C727AAA-F395-40EC-990A-36B7D7A73A64}"
Private Const DETAILPROFILEPARTINPUT_DEST = "DetailProfilePartInput_DEST"
Private Const MFGPROFILEOUTPUT_ORIG = "MfgProfileOutput_ORIG"
Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"

'Localizer strings
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
' Abstract: Implemented by IJCheckMfcty interface
' Description:
' This method is called by the Check Manufacturability mechanism
' to perform the desired checks for this rule.This method returns
' the profile parts which have more than one manufacturing profile
' Inputs:
'     oElements     :   The selected items in the working set.
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
    
    'Setup the Check Manufacturability call back mechanism.
    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    'If the working set is empty, do not continue.
    If oElements.Count = 0 Then
        Exit Sub
    End If
    
    InitializeLocalizeVariables
    
    'Get the Pesistent Object Manager.
    Set oPOM = GetPOMEx("Model")
       
    'Loop
    nCount = 0
    nMaxCount = oElements.Count
    
    m_bLogError = False
    For Each vMoniker In oElements
        'Get the object from the moniker.
        Set oObject = oPOM.GetObject(vMoniker)
        
        'Check the values.
        bErrorOccurred = CheckObjectForErrorCondition(oObject, strMessage)
        
        If bErrorOccurred Then
            
            
            'The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oObject                the object with error/warning
            '   strProgID              prog id of this rule
            '   ESeverity.siError      severity of the error
            '                          (101 error or 102 warning)
            '   106(2)                    code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            '   strAction              type of repair to be made
            '   strActionTool          progid of the COM object that will make the repair
            oCallBack.OnCheckError oObject, strProgID, _
                                   ESeverityIndex.siError, 2, strMessage, _
                                   "COM Repair Object", _
                                   "StrMfgCheckMfctyProfile.RepairProfile"
            
            'The operator may desire that only the first error be found.
            If bStopOnError Then
                Exit For
            End If
        End If
            
        'Update the progress bar.  First argument is the progress as a
        'percentage (i.e. 0 - 100).  The second argument is a cancellation
        'flag.
        nCount = nCount + 1
        oCallBack.OnProgress nCount * 100 / nMaxCount, 0
    Next vMoniker

    If m_bLogError Then
        MsgBox m_ProcessingCompleted & vbCrLf & _
               m_ErrDetected & vbCrLf & _
               m_ViewLogFile, , _
               "StrMfgCheckMfctyProfile.CCheckProfile"
    End If
    
    Set oCallBack = Nothing
    Set vMoniker = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
End Sub

'******************************************************************************
' Routine: CheckObjectForErrorCondition
' Abstract:
'  Checks if the current Object is a Profile part
'  If it is a profile part,get the number of manufacturing profiles it has.
'  If there is a concurrency,report it.
' Description:
' Inputs:
' Outputs:
'******************************************************************************
Private Function CheckObjectForErrorCondition(oObject As Object, _
                                              ByRef errorMessage As String) As Boolean
    Const METHOD As String = "CheckObjectForErrorCondition"
    On Error GoTo ErrorHandler
       
    Dim oUnk As IUnknown
    Dim sMessage As String
    
    Dim oMfgUpdateService  As IJMfgUpdateService
    Set oMfgUpdateService = New MfgUpdateService
    
    CheckObjectForErrorCondition = False
        
    'Check if current object is type IJStructProfilePart
    If TypeOf oObject Is IJStructProfilePart Then
        
       Dim oAssocRel As IJDAssocRelation
       Dim oUnkColl As Object
       Dim oIJDObject As IJDObject
       Dim oTargetObjCol As IJDTargetObjectCol

       Set oAssocRel = oObject 'Profile part

       Set oUnkColl = oAssocRel.CollectionRelations(IID_IJStructProfilePart, DETAILPROFILEPARTINPUT_DEST)
       Set oTargetObjCol = oUnkColl
       If oTargetObjCol.Count > 1 Then
            'If there is more than one Manufacturing profile then check if it is a result of partial split migration
             Set oAssocRel = oTargetObjCol.Item(1) 'AE
             Set oTargetObjCol = oAssocRel.CollectionRelations _
                                                    (IID_IJMfgProfileCreation_AE, MFGPROFILEOUTPUT_ORIG)
                                                    
             Set oIJDObject = oTargetObjCol.Item(1) 'Mfg profile
             
             oMfgUpdateService.CheckForPartialUpdate oIJDObject, oUnk
             
             If oUnk Is Nothing Then 'Mfg profiles are not as a result of partial split migration
                CheckObjectForErrorCondition = True
                errorMessage = m_ErrMsg
             End If
        End If
    End If
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
    
    m_ErrMsg = oLocalizer.GetString(IDS_CHECKMFGPROFILE_MSG_ERRORMESSAGE, "The profile part has invalid number of Manufacturing profiles")
    m_ViewLogFile = oLocalizer.GetString(IDS_CHECKMFGPROFILE_MSG_VIEWLOGFILE, "See Sp3d ErrorLog for Manufacturing")
    m_ProcessingCompleted = oLocalizer.GetString(IDS_CHECKMFGPROFILE_MSG_PROCESSINGCOMPLETE, "Processing Complete")
    m_ErrDetected = oLocalizer.GetString(IDS_CHECKMFGPROFILE_MSG_ERRDETECTED, "Errors (and/or Warnings) were detected")
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Sub
