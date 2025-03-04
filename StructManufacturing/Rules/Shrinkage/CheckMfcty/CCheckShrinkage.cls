VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckShrinkage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
' File: CCheckShrinkage.cls
' Author: J.Manasa
' Abstract:
'   Check Manufacturability rule that checks the concurrency of Shrinkages
'   This rule checks if there is more than one Shrinkage for a part
'********************************************************************
Option Explicit

Private Const Module = "StrMfgCheckMfctyShr.CCheckShrinkage"
Private Const strProgID = "StrMfgCheckMfctyShr.CCheckShrinkage"

Private Const IID_IJGeometry = "{96EB9676-6530-11D1-977F-080036754203}"
Private Const IID_IJShrinkage_AE = "{5E839CD8-EEEF-45E3-9339-06075272AA2E}"
Private Const INPUTGEOMETRY_AE = "InputGeometry_AE"
Private Const SHRINKAGEOUTPUT_SHRINKAGE = "ShrinkageOutput_Shrinkage"
Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"
Private Const IID_IJScalingShr = "{DE77050C-3300-11D5-BA1A-0090276F4279}"

'Localizer strings
Private m_ErrPlate As String
Private m_ErrProfile As String
Private m_ErrAssembly As String
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
' the parts which have more than one shrinkage
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
            '   106(3)                    code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            '   strAction              type of repair to be made
            '   strActionTool          progid of the COM object that will make the repair
            oCallBack.OnCheckError oObject, strProgID, _
                                   ESeverityIndex.siError, 3, strMessage, _
                                   "COM Repair Object", _
                                   "StrMfgCheckMfctyShr.RepairShrinkage"
            
            'The operator may desire that only the first error be found.
            If bStopOnError Then
                Exit For
            End If
        End If
            
        'Update the progress bar. First argument is the progress as a
        'percentage (i.e. 0 - 100). The second argument is a cancellation
        'flag.
        nCount = nCount + 1
        oCallBack.OnProgress nCount * 100 / nMaxCount, 0
        
    Next vMoniker

    If m_bLogError Then
        MsgBox m_ProcessingCompleted & vbCrLf & _
               m_ErrDetected & vbCrLf & _
               m_ViewLogFile, , _
               "StrMfgCheckMfctyShr.CCheckShrinkage"
               
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
'  Checks if the current Object is a Plate or Profile part
'  If it is a plate or profile part,get the number of shrinkages it has.
'  If there is a concurrency and it is not as a result of partial migration report it.
' Description:
' Inputs:
' Outputs:
'******************************************************************************
Private Function CheckObjectForErrorCondition(oObject As Object, _
                                              ByRef errorMessage As String) As Boolean
    Const METHOD As String = "CheckObjectForErrorCondition"
    On Error GoTo ErrorHandler

    Dim sMessage As String
    
    CheckObjectForErrorCondition = False
        
    'Check if current object is type of IJSturctProfilePart/IJPlatePart/IJAssembly
    If TypeOf oObject Is IJStructProfilePart Or TypeOf oObject Is IJPlatePart Or TypeOf oObject Is IJAssembly Then

        Dim oAllShrinkages As IJElements
        Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
       
        Set oAllShrinkages = oStructMfgGlobals.GetMfgPart(oObject, IID_IJScalingShr)
        
        If Not oAllShrinkages Is Nothing Then
            If oAllShrinkages.Count > 1 Then
                Dim i As Integer
                Dim j As Integer
                
                For i = 1 To oAllShrinkages.Count ' this loop checks whether shrinkage modes are same for different shrinkages
                    For j = i + 1 To oAllShrinkages.Count
                        Dim oShrinkage1 As IJScalingShr
                        Dim oShrinkage2 As IJScalingShr
                        
                        Set oShrinkage1 = oAllShrinkages(i)
                        Set oShrinkage2 = oAllShrinkages(j)
                        
                        If oShrinkage1.ShrinkageMode = oShrinkage2.ShrinkageMode Then
                            CheckObjectForErrorCondition = True
                            
                            If TypeOf oObject Is IJPlatePart Then
                               errorMessage = m_ErrPlate
                            ElseIf TypeOf oObject Is IJAssembly Then
                               errorMessage = m_ErrAssembly
                            Else
                               errorMessage = m_ErrProfile
                            End If

                            Exit Function
                        End If
                    Next
                Next
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
    
    m_ErrPlate = oLocalizer.GetString(IDS_CHECKSHRINKAGE_MSG_PLATEERRMSG, "The plate part has invalid number of Shrinkages")
    m_ErrProfile = oLocalizer.GetString(IDS_CHECKSHRINKAGE_MSG_PROFILEERRMSG, "The profile part has invalid number of Shrinkages")
    m_ErrAssembly = oLocalizer.GetString(IDS_CHECKSHRINKAGE_MSG_ASSEMBLYFILEERRMSG, "The Assembly has invalid number of Shrinkages")
    m_ViewLogFile = oLocalizer.GetString(IDS_CHECKSHRINKAGE_MSG_VIEWLOGFILE, "See Sp3d ErrorLog for Manufacturing")
    m_ProcessingCompleted = oLocalizer.GetString(IDS_CHECKSHRINKAGE_MSG_PROCESSINGCOMPLETE, "Processing Complete")
    m_ErrDetected = oLocalizer.GetString(IDS_CHECKSHRINKAGE_MSG_ERRDETECTED, "Errors (and/or Warnings) were detected")
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Sub
