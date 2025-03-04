VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
' File: CCheckMargin.cls
' Author: J.Manasa
' Abstract:
'   Check Manufacturability rule that checks the concurrency of margins of same type on a particular port
'********************************************************************
Option Explicit

Private Const Module = "StrMfgCheckMfctyMargin.CCheckMargin"
Private Const strProgID = "StrMfgCheckMfctyMargin.CCheckMargin"

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
' the parts which have more than one margin of same type on a given port
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
    'Dim Index As Integer
    'Set m_MarginColl = New elements
   
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
            
            'For Index = 1 To m_MarginColl.Count
            'The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oObject                the object with error/warning
            '   strProgID              prog id of this rule
            '   ESeverity.siError      severity of the error
            '                          (101 error or 102 warning)
            '   106(4)                    code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            '   strAction              type of repair to be made
            '   strActionTool          progid of the COM object that will make the repair
            oCallBack.OnCheckError oObject, strProgID, _
                                   ESeverityIndex.siError, 4, strMessage, _
                                   "COM Repair Object", _
                                   "StrMfgCheckMfctyMargin.RepairMargin"
            'Next
            'm_MarginColl.Clear
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
               "StrMfgCheckMfctyMargin.CCheckMargin"
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
'  Checks if the current Object is a plate or profile part
'  If it is a plate part,get the number of margins of same type on a given port.
'  If there is a concurrency and it is not as a result of partial split then report it.
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
    Dim eObject As Object
    
    Dim oPort As IJPort
    Dim oEligiblePort As IJPort
    
    Dim oStructConnectable  As IJStructConnectable
    Dim oMfgDefCol As Collection
    Dim oEnumPorts As IMSCoreCollections.IJElements
    Dim m_oEligiblePortColl As IJDObjectCollection
    
    Dim oMargin As IJDFabMargin
    Dim oOtherMargin As IJDFabMargin
    
    Dim i As Integer
    Dim n As Integer
    Dim j As Integer
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    Dim oMfgUpdateService  As IJMfgUpdateService
    Set oMfgUpdateService = New MfgUpdateService
    
    CheckObjectForErrorCondition = False
    
    Set oEnumPorts = New JObjectCollection
    'Check if current object is of type IJPlatePart or IJStructProfilePart
         If TypeOf oObject Is IJPlatePart Or TypeOf oObject Is IJProfilePart Then
            Set oStructConnectable = oObject
            'Get all the ports
            sMessage = "...Getting all the ports"
            oEnumPorts.AddElements oStructConnectable.enumAllPorts
            Set oStructConnectable = Nothing
         ElseIf TypeOf oObject Is ISPSMemberPartPrismatic Then
            Dim oPartSupport As IJPartSupport
            Set oPartSupport = New MemberPartSupport
            Set oPartSupport.Part = oObject
            Set oStructConnectable = oPartSupport
            oEnumPorts.AddElements oStructConnectable.enumAllPorts
            Set oStructConnectable = Nothing
            Set oPartSupport = Nothing
        Else
            Exit Function
        End If
        
        Set m_oEligiblePortColl = New JObjectCollection
        
        'Get the eligible port collection
        sMessage = "...Getting the eligible port collection"

        For j = 1 To oEnumPorts.Count
            Set oEligiblePort = oEnumPorts.Item(j)
            If TypeOf oObject Is IJPlatePart Then
                If IsLateralFacePort(oEligiblePort) Then
                    m_oEligiblePortColl.Add oEligiblePort
                    Set oEligiblePort = Nothing
                End If
            ElseIf TypeOf oObject Is IJStructProfilePart Then
                If IsBaseOrOffsetFacePort(oEligiblePort) Then
                    m_oEligiblePortColl.Add oEligiblePort
                    Set oEligiblePort = Nothing
                End If
            End If
        Next j
        
        sMessage = "...Iterating through the eligible port collection"
        
        'Iterate through all the ports
        For Each eObject In m_oEligiblePortColl
            Set oPort = eObject
                            
            'Get all the Mfgdefinitions on that port
            Set oMfgDefCol = oMfgRuleHelper.GetMfgDefinitions(oPort)
            
            For i = 1 To oMfgDefCol.Count
                If (TypeOf oMfgDefCol.Item(i) Is IJDFabMargin) Then
                    Set oMargin = oMfgDefCol.Item(i)
                    For j = 1 To oMfgDefCol.Count
                        Set oOtherMargin = oMfgDefCol.Item(j)
                        If Not oOtherMargin Is oMargin Then
                            If CInt(oMargin.MarginType) = CInt(oOtherMargin.MarginType) And CInt(oMargin.MarginMode) = CInt(oOtherMargin.MarginMode) Then
                                oMfgUpdateService.CheckForPartialUpdate oMargin, oUnk
                                If oUnk Is Nothing Then  'Margins are not as a result of partial split migration
                                     CheckObjectForErrorCondition = True
                                     errorMessage = m_ErrMsg
                                     Exit For
                                End If
                            End If
                        End If
                    Next 'oMfgDefCol j
                End If
            Next 'oMfgDefCol i
        Next 'm_oEligiblePortColl
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
    
    m_ErrMsg = oLocalizer.GetString(IDS_CHECKMARGIN_MSG_ERRMSG, "There are invalid number of margins of same type on a given port")
    m_ViewLogFile = oLocalizer.GetString(IDS_CHECKMARGIN_MSG_VIEWLOGFILE, "See Sp3d ErrorLog for Manufacturing")
    m_ProcessingCompleted = oLocalizer.GetString(IDS_CHECKMARGIN_MSG_PROCESSINGCOMPLETE, "Processing Complete")
    m_ErrDetected = oLocalizer.GetString(IDS_CHECKMARGIN_MSG_ERRDETECTED, "Errors (and/or Warnings) were detected")
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Sub

Public Function IsLateralFacePort(pObj As Object) As Long
    Const METHOD As String = "IsLateralFacePort"
    On Error GoTo ErrorHandler
   
    IsLateralFacePort = 0

    '- Return if it not a face port
    If TypeOf pObj Is IJPort Then
        Dim oPort As IJPort
        Set oPort = pObj
        If oPort.Type <> PortFace Then
            Set oPort = Nothing
            Exit Function
        End If
        Set oPort = Nothing
    End If

    '- Check for lateral port
    If Not pObj Is Nothing Then
        If TypeOf pObj Is IJStructPort Then
            Dim oStructPort As IJStructPort
            Set oStructPort = pObj
            If oStructPort.ContextID = CTX_LATERAL_LFACE Then 'Or oStructPort.ContextID = CTX_LATERAL_LFACE_3DSPLIT Then
                IsLateralFacePort = 1
            End If
            Set oStructPort = Nothing
        End If
    End If
    Exit Function
ErrorHandler:
    IsLateralFacePort = 0
End Function

Public Function IsBaseOrOffsetFacePort(pObj As Object) As Long
    Const METHOD As String = "IsBaseOrOffsetFacePort"
    On Error GoTo ErrorHandler
   
    Dim sMessage As String
    IsBaseOrOffsetFacePort = 0
         
    '- Check for lateral port
    sMessage = "...Checking the ContextID"
    If Not pObj Is Nothing Then
        If TypeOf pObj Is IJStructPort Then
            Dim oStructPort As IJStructPort
            Set oStructPort = pObj
            If oStructPort.ContextID = CTX_BASE Or oStructPort.ContextID = CTX_OFFSET Then
                IsBaseOrOffsetFacePort = 1
            End If
            Set oStructPort = Nothing
        End If
    End If
    Exit Function
ErrorHandler:
    IsBaseOrOffsetFacePort = 0
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Function
